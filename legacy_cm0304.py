"""
CM03/04 legacy parser boundary.

The adapter in app.py performs safe inventory of a CM03/04 database folder.
This module is intentionally a separate boundary so the actual binary parser
can be dropped in without changing the rest of the workbench.

A future validated implementation should use the record structures from the
archived Nygreen CM4/03-04 editor source and be tested against:
- stock 4.1.4 / 4.1.5 CM03/04 database
- at least one modified database
- checksum/round-trip samples
- counts for people/clubs/nations/competitions/history tables
"""

from dataclasses import dataclass
from pathlib import Path

@dataclass
class CM0304ParseResult:
    players: list
    staff: list
    clubs: list
    nations: list
    competitions: list
    transfers: list
    histories: list

class CM0304BinaryParser:
    def __init__(self, folder: Path):
        self.folder = Path(folder)

    def parse(self) -> CM0304ParseResult:
        raise NotImplementedError(
            "Validated CM03/04 binary decoding is the next parser milestone. "
            "The workbench deliberately refuses to guess at proprietary record layouts."
        )
