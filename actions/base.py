from abc import ABC, abstractmethod
from typing import List, Dict, Any

class TransferAction(ABC):
    """Abstract base class for transfer actions."""
    
    @abstractmethod
    def analyze(self) -> List[Any]:
        """Analyzes the transfer and returns diffs/preview data."""
        pass
    
    @abstractmethod
    def execute(self, diffs: List[Any], **kwargs):
        """Executes the transfer based on diffs."""
        pass
