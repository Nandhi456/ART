from typing import Optional, List, Dict, Any
from pydantic import BaseModel


class ExtractRequest(BaseModel):
    folder_name: str  
    destination_name: Optional[str] = None  


class ExtractResponse(BaseModel):
    folder_name: str
    file_count: int


class FolderFile(BaseModel):
    filename: str
    type: str
    size_bytes: int


class Statistics(BaseModel):
    job_id: str
    status: str  # "pending" | "processing" | "done" | "error"
    processed: int
    total: int
    error: Optional[str] = None


class PreviewData(BaseModel):
    folder_name: str
    columns: List[str]
    rows: List[Dict[str, Any]]
    total_files: int
    raw_records: int
    cleaned_records: int
    removed: int
    failed_files: int


class SearchRequest(BaseModel):
    keyword: str

class RecentFiles(BaseModel):
    Filename:str
