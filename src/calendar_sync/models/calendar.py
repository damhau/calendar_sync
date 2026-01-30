"""Calendar metadata model."""

from typing import Optional

from pydantic import BaseModel


class Calendar(BaseModel):
    """Calendar metadata."""

    id: str
    name: str
    owner_email: Optional[str] = None
    source_system: str  # "ews" or "m365"
    is_default: bool = False
    can_edit: bool = False
    color: Optional[str] = None

    model_config = {"frozen": True}
