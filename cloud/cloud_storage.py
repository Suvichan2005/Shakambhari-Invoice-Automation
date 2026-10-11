"""
Google Cloud Storage Integration for Shakambhari Invoice Generator
===================================================================
Handles file uploads and downloads for:
- Generated Excel invoices
- Generated PDF invoices
- Invoice templates
"""
from __future__ import annotations
import os
import io
from typing import Optional, Tuple
from datetime import datetime, timedelta
try:
    from google.cloud import storage
    from google.cloud.exceptions import NotFound
except ImportError:
    storage = None
    NotFound = Exception


class CloudStorage:
    """
    Interface for Google Cloud Storage operations.
    """
    
    def __init__(self, bucket_name: str = None):
        """
        Initialize Cloud Storage connection.
        
        Args:
            bucket_name: Name of the GCS bucket
        """
        self.bucket_name = bucket_name or os.environ.get('GCS_BUCKET_NAME')
        self.client = storage.Client()
        self.bucket = self.client.bucket(self.bucket_name)
        
        # Folder structure in bucket
        self.INVOICES_FOLDER = 'invoices/'
        self.PDFS_FOLDER = 'pdfs/'
        self.TEMPLATES_FOLDER = 'templates/'
    
    def upload_file(self, file_data: bytes, destination_path: str, content_type: str = None) -> str:
        """
        Upload a file to Cloud Storage.
        
        Args:
            file_data: File content as bytes
            destination_path: Path in the bucket (e.g., 'invoices/Invoice_123.xlsx')
            content_type: MIME type of the file
        
        Returns:
            Public URL of the uploaded file
        """
        blob = self.bucket.blob(destination_path)

        # Pass content type in the same upload call so metadata and payload match.
        if content_type:
            blob.upload_from_string(file_data, content_type=content_type)
        else:
            blob.upload_from_string(file_data)
        
        # Make the file publicly readable (or use signed URLs for private access)
        # blob.make_public()
        
        return f"gs://{self.bucket_name}/{destination_path}"
    
    def upload_invoice_xlsx(self, file_data: bytes, filename: str) -> str:
        """Upload an Excel invoice."""
        path = f"{self.INVOICES_FOLDER}{filename}"
        return self.upload_file(
            file_data, 
            path, 
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    def upload_invoice_pdf(self, file_data: bytes, filename: str) -> str:
        """Upload a PDF invoice."""
        path = f"{self.PDFS_FOLDER}{filename}"
        return self.upload_file(file_data, path, content_type='application/pdf')
    
    def download_file(self, source_path: str) -> Optional[bytes]:
        """
        Download a file from Cloud Storage.
        
        Args:
            source_path: Path in the bucket
        
        Returns:
            File content as bytes, or None if not found
        """
        try:
            blob = self.bucket.blob(source_path)
            return blob.download_as_bytes()
        except NotFound:
            return None
    
    def download_invoice_xlsx(self, filename: str) -> Optional[bytes]:
        """Download an Excel invoice."""
        return self.download_file(f"{self.INVOICES_FOLDER}{filename}")
    
    def download_invoice_pdf(self, filename: str) -> Optional[bytes]:
        """Download a PDF invoice."""
        return self.download_file(f"{self.PDFS_FOLDER}{filename}")
    
    def get_signed_url(self, source_path: str, expiration_minutes: int = 60) -> str:
        """
        Generate a signed URL for temporary access to a file.
        
        Args:
            source_path: Path in the bucket
            expiration_minutes: How long the URL should be valid
        
        Returns:
            Signed URL string
        """
        blob = self.bucket.blob(source_path)
        url = blob.generate_signed_url(
            expiration=timedelta(minutes=expiration_minutes),
            method='GET'
        )
        return url
    
    def get_invoice_download_url(self, filename: str, file_type: str = 'xlsx', expiration_minutes: int = 60) -> str:
        """
        Get a download URL for an invoice.
        
        Args:
            filename: Invoice filename
            file_type: 'xlsx' or 'pdf'
            expiration_minutes: URL validity period
        
        Returns:
            Signed download URL
        """
        if file_type == 'pdf':
            path = f"{self.PDFS_FOLDER}{filename}"
        else:
            path = f"{self.INVOICES_FOLDER}{filename}"
        
        return self.get_signed_url(path, expiration_minutes)
    
    def list_invoices(self, limit: int = 100) -> list:
        """
        List all invoice files in storage.
        
        Returns:
            List of invoice file info dicts
        """
        invoices = []
        blobs = self.client.list_blobs(self.bucket_name, prefix=self.INVOICES_FOLDER)
        
        for blob in blobs:
            if blob.name.endswith('.xlsx'):
                filename = blob.name.replace(self.INVOICES_FOLDER, '')
                invoices.append({
                    'filename': filename,
                    'path': blob.name,
                    'size': blob.size,
                    'created': blob.time_created,
                    'updated': blob.updated
                })
        
        # Sort by updated time descending
        invoices.sort(key=lambda x: x['updated'], reverse=True)
        return invoices[:limit]
    
    def delete_file(self, source_path: str) -> bool:
        """Delete a file from storage."""
        try:
            blob = self.bucket.blob(source_path)
            blob.delete()
            return True
        except NotFound:
            return False
    
    def upload_template(self, file_data: bytes, filename: str) -> str:
        """Upload an invoice template."""
        path = f"{self.TEMPLATES_FOLDER}{filename}"
        return self.upload_file(
            file_data,
            path,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    def download_template(self, filename: str = None) -> Optional[Tuple[bytes, str]]:
        """
        Download the invoice template.
        
        Returns:
            Tuple of (file_data, filename) or None if not found
        """
        # If no specific filename, get the first template
        if not filename:
            blobs = list(self.client.list_blobs(self.bucket_name, prefix=self.TEMPLATES_FOLDER))
            xlsx_blobs = [b for b in blobs if b.name.endswith('.xlsx')]
            if xlsx_blobs:
                blob = xlsx_blobs[0]
                return (blob.download_as_bytes(), blob.name.replace(self.TEMPLATES_FOLDER, ''))
            return None
        
        data = self.download_file(f"{self.TEMPLATES_FOLDER}{filename}")
        if data:
            return (data, filename)
        return None


def init_cloud_storage() -> CloudStorage:
    """Initialize and return the Cloud Storage connection."""
    bucket_name = os.environ.get('GCS_BUCKET_NAME')
    if not bucket_name:
        raise ValueError("GCS_BUCKET_NAME environment variable is not set")
    
    return CloudStorage(bucket_name=bucket_name)


# For testing locally
if __name__ == '__main__':
    import os
    os.environ['GCS_BUCKET_NAME'] = 'shakambhari-invoices-bucket'
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'service-account.json'
    
    storage = init_cloud_storage()
    print("Invoices in storage:", storage.list_invoices(limit=5))
