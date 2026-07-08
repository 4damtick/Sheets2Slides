import base64
import json
import logging
import os
from io import BytesIO
from urllib.parse import unquote

import fitz  # PyMuPDF
import numpy as np
from fastapi import FastAPI, HTTPException, Request
from google.api_core.exceptions import NotFound
from google.cloud import storage
from PIL import Image

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

app = FastAPI(title="PDF to PNG Converter Worker")


@app.get("/")
async def root():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "service": "PDF to PNG Converter",
        "mode": "pubsub-worker",
        "version": "3.0.0",
        "engine": "PyMuPDF",
        "features": [
            "GCS input",
            "Pub/Sub trigger",
            "GCS output",
            "Auto-crop margins",
            "Manual crop",
            "Transparent background"
        ]
    }

def auto_crop_image(image, threshold=250, padding=10):
    """
    Detect non-white pixels with threshold
    More aggressive cropping based on actual pixel values
    """
    # Convert to numpy array for faster processing
    img_array = np.array(image)
    
    # Convert to grayscale if RGB/RGBA
    if len(img_array.shape) == 3:
        # Consider a pixel "white" if all channels are above threshold
        mask = np.any(img_array[:, :, :3] < threshold, axis=2)
    else:
        mask = img_array < threshold
    
    # Find rows and columns with content
    rows = np.any(mask, axis=1)
    cols = np.any(mask, axis=0)
    
    # Find the bounding box
    if not rows.any() or not cols.any():
        return image  # No content found, return original
    
    row_indices = np.where(rows)[0]
    col_indices = np.where(cols)[0]
    
    top = max(0, row_indices[0] - padding)
    bottom = min(img_array.shape[0], row_indices[-1] + 1 + padding)
    left = max(0, col_indices[0] - padding)
    right = min(img_array.shape[1], col_indices[-1] + 1 + padding)
    
    return image.crop((left, top, right, bottom))

def manual_crop_image(image, crop_left, crop_top, crop_right, crop_bottom):
    """
    Manually crop image using pixel coordinates
    """
    width, height = image.size
    
    # Calculate crop box (left, top, right, bottom)
    left = crop_left if crop_left >= 0 else 0
    top = crop_top if crop_top >= 0 else 0
    right = width - crop_right if crop_right >= 0 else width
    bottom = height - crop_bottom if crop_bottom >= 0 else height
    
    # Ensure valid crop box
    if left >= right or top >= bottom:
        raise ValueError("Invalid crop coordinates")
    
    return image.crop((left, top, right, bottom))


def _clamp_int(value, default, min_value, max_value):
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return default
    return max(min_value, min(max_value, parsed))


def _parse_bool(value, default=True):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _build_output_name(input_name):
    if input_name.lower().endswith(".pdf"):
        return input_name[:-4] + ".png"
    return input_name + ".png"


def convert_pdf_bytes_to_png(
    pdf_bytes,
    dpi=300,
    quality=95,
    background="white",
    page=1,
    auto_crop=True,
    crop_left=0,
    crop_top=0,
    crop_right=0,
    crop_bottom=0,
):
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        if page < 1 or page > len(pdf_document):
            raise ValueError(f"Page {page} does not exist. PDF has {len(pdf_document)} page(s)")

        pdf_page = pdf_document[page - 1]
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        use_alpha = str(background).lower() in ["transparent", "none"]

        pix = pdf_page.get_pixmap(matrix=mat, alpha=use_alpha)
        png_bytes = pix.tobytes("png")
        img = Image.open(BytesIO(png_bytes))

        bg_lower = str(background).lower()
        if bg_lower in ["transparent", "none"]:
            img = img.convert("RGBA")
        elif bg_lower not in ["white", "#ffffff", "#fff"]:
            if img.mode == "RGBA":
                bg_image = Image.new("RGB", img.size, background)
                bg_image.paste(img, (0, 0), img)
                img = bg_image
            else:
                bg_image = Image.new("RGB", img.size, background)
                bg_image.paste(img, (0, 0))
                img = bg_image

        if auto_crop:
            img = auto_crop_image(img)

        if crop_left > 0 or crop_top > 0 or crop_right > 0 or crop_bottom > 0:
            img = manual_crop_image(img, crop_left, crop_top, crop_right, crop_bottom)

        img_byte_arr = BytesIO()
        compress_level = int((100 - quality) / 100 * 9)
        img.save(img_byte_arr, format="PNG", optimize=True, compress_level=compress_level)
        img_byte_arr.seek(0)
        return img_byte_arr.getvalue()
    finally:
        pdf_document.close()


@app.post("/events/pdf-uploaded")
async def handle_pubsub_pdf_uploaded(request: Request):
    """Handle Pub/Sub push event produced by Cloud Storage notifications."""
    try:
        envelope = await request.json()
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Invalid JSON payload: {exc}")

    message = envelope.get("message")
    if not message:
        raise HTTPException(status_code=400, detail="Missing Pub/Sub message envelope")

    encoded_data = message.get("data")
    if not encoded_data:
        raise HTTPException(status_code=400, detail="Missing Pub/Sub message data")

    try:
        decoded = base64.b64decode(encoded_data).decode("utf-8")
        event_data = json.loads(decoded)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Failed to decode Pub/Sub data: {exc}")

    bucket_name = event_data.get("bucketId") or event_data.get("bucket")
    object_name = unquote(event_data.get("objectId") or event_data.get("name") or "")

    if not bucket_name or not object_name:
        raise HTTPException(status_code=400, detail="Missing bucket or object in event payload")

    if not object_name.lower().endswith(".pdf"):
        logger.info("Ignoring non-PDF object: gs://%s/%s", bucket_name, object_name)
        return {
            "status": "ignored",
            "reason": "non-pdf object",
            "bucket": bucket_name,
            "object": object_name,
        }

    output_bucket_name = os.environ.get("BUCKET_OUTPUT", bucket_name)
    logger.info("Processing gs://%s/%s → gs://%s", bucket_name, object_name, output_bucket_name)

    try:
        storage_client = storage.Client()
        input_bucket = storage_client.bucket(bucket_name)
        pdf_blob = input_bucket.blob(object_name)

        try:
            pdf_bytes = pdf_blob.download_as_bytes()
        except NotFound:
            # The PDF was deleted by the client before this (retry) delivery
            # arrived.  Returning 200 tells Pub/Sub the message is acknowledged
            # so it stops retrying.
            logger.info(
                "PDF already deleted — skipping (duplicate/retry delivery): gs://%s/%s",
                bucket_name, object_name,
            )
            return {
                "status": "skipped",
                "reason": "pdf_already_deleted",
                "input": f"gs://{bucket_name}/{object_name}",
            }

        logger.info("Downloaded PDF (%d bytes)", len(pdf_bytes))

        metadata = pdf_blob.metadata or {}
        dpi = _clamp_int(metadata.get("dpi"), default=300, min_value=72, max_value=600)
        quality = _clamp_int(metadata.get("quality"), default=95, min_value=1, max_value=100)
        page = _clamp_int(metadata.get("page"), default=1, min_value=1, max_value=5000)
        background = metadata.get("background", "white")
        auto_crop = _parse_bool(metadata.get("auto_crop"), default=True)
        crop_left = _clamp_int(metadata.get("crop_left"), default=0, min_value=0, max_value=10000)
        crop_top = _clamp_int(metadata.get("crop_top"), default=0, min_value=0, max_value=10000)
        crop_right = _clamp_int(metadata.get("crop_right"), default=0, min_value=0, max_value=10000)
        crop_bottom = _clamp_int(metadata.get("crop_bottom"), default=0, min_value=0, max_value=10000)

        output_name = _build_output_name(object_name)
        png_bytes = convert_pdf_bytes_to_png(
            pdf_bytes=pdf_bytes,
            dpi=dpi,
            quality=quality,
            background=background,
            page=page,
            auto_crop=auto_crop,
            crop_left=crop_left,
            crop_top=crop_top,
            crop_right=crop_right,
            crop_bottom=crop_bottom,
        )

        output_bucket = storage_client.bucket(output_bucket_name)
        output_blob = output_bucket.blob(output_name)
        output_blob.metadata = {
            "source_pdf": f"gs://{bucket_name}/{object_name}",
            "dpi": str(dpi),
            "quality": str(quality),
            "page": str(page),
        }
        output_blob.upload_from_string(png_bytes, content_type="image/png")
        logger.info("Uploaded PNG to gs://%s/%s (%d bytes)", output_bucket_name, output_name, len(png_bytes))

        return {
            "status": "processed",
            "input": f"gs://{bucket_name}/{object_name}",
            "output": f"gs://{output_bucket_name}/{output_name}",
        }
    except HTTPException:
        raise
    except Exception as exc:
        logger.error("Failed to process gs://%s/%s: %s", bucket_name, object_name, exc)
        raise HTTPException(status_code=500, detail=f"Failed to process Pub/Sub event: {exc}")

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(app, host="0.0.0.0", port=port)
