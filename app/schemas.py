# schemas.py
# Defines data structures for personnel and calendar information

from dataclasses import dataclass
from typing import Optional

@dataclass
class Person:
    rank: str
    name: str
    specialty: str
    duty: str
    primary_shift: Optional[str] = None
    alt_shift: Optional[str] = None
    at_sea_shift: Optional[str] = None
    height: Optional[float] = None
    weight: Optional[float] = None
    registry_number: Optional[str] = None
    address: Optional[str] = None
    phone: Optional[str] = None
    marital_status: Optional[str] = None
    children: Optional[int] = None
    pye_expiration: Optional[str] = None
    notes: Optional[str] = None

@dataclass
class LeaveRecord:
    person_name: str
    leave_type: str
    start_date: str
    end_date: str
    comments: Optional[str] = None

@dataclass
class AvailabilityRecord:
    person_name: str
    unavailable_dates: str  # e.g., "5,8,12"
    comments: Optional[str] = None

@dataclass
class Holiday:
    date: str
    description: str

@dataclass
class ShipStatus:
    date: str
    status: str  # "In Port" or "At Sea"
