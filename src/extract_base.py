from dataclasses import dataclass
from abc import ABC, abstractmethod
from typing import Iterator


@dataclass(frozen=True)
class Record:
    """One row from Excel with only the important values."""
    sample_id: str              # Column G
    sample_type: str            # Column H 
    mean: float | None          # Column I
    ppm: float | None           # Column L
    adjusted_abs: float | None  # Column M -> Needed for "Calibration -> Will work on it soon



class Extractor(ABC):
    @abstractmethod
    def records(self) -> Iterator[Record]:
        """Give back one Record at a time."""
        raise NotImplementedError