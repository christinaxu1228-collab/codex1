"""SHEIN OpenAPI contract constants; no network calls or secrets here.

Endpoints must be checked against the shop's actual authorization and latest docs.
"""
from dataclasses import dataclass

@dataclass(frozen=True)
class SheinEndpoints:
    publish: str = "/open-api/goods/product/publishOrEdit"
    publish_requirements: str = "/open-api/goods/query-publish-fill-in-standard"
    category_tree: str = "/open-api/goods/query-category-tree"
    attributes: str = "/open-api/goods/query-attribute-template"
    image_upload: str = "/open-api/goods/upload-pic"
    image_transform: str = "/open-api/goods/transform-pic"
    publish_permission: str = "/open-api/goods/product/check-publish-permission"
    review_status: str = "/open-api/goods/query-document-state"
    feed_document_create: str = "/open-api/sem/feed/createFeedDocument"
    feed_document_upload: str = "/open-api/sem/feed/uploadDocumentContent"
    feed_document_get: str = "/open-api/sem/feed/getFeedDocument"
    feed_create: str = "/open-api/sem/feed/createFeed"
    feed_get: str = "/open-api/sem/feed/getFeed"


def feed_task_request(feed_document_id: str) -> dict:
    """Generate known documented fields; does NOT submit anything."""
    if not feed_document_id.strip():
        raise ValueError("feed_document_id is required")
    return {
        "feedDocumentId": feed_document_id,
        "feedType": "PRODUCT_LISTING",
        "version": "V2",
    }
