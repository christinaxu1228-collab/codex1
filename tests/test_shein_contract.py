import unittest
from bulk_listing.shein_contract import SheinEndpoints, feed_task_request

class SheinContractTests(unittest.TestCase):
    def test_feed_request(self):
        self.assertEqual(feed_task_request("demo.json"), {
            "feedDocumentId": "demo.json", "feedType": "PRODUCT_LISTING", "version": "V2"
        })
    def test_empty_document_rejected(self):
        with self.assertRaises(ValueError):
            feed_task_request("   ")
    def test_publish_endpoint(self):
        self.assertEqual(SheinEndpoints().publish, "/open-api/goods/product/publishOrEdit")
