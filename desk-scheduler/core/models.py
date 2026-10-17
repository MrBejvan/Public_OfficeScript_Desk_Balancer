"""
Core data models for the Desk Coverage Scheduler.

These are plain dataclasses with no external dependencies so they can be
unit-tested in isolation from the GUI and from file I/O.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from enum import Enum


class Desk(str, Enum):
    """The four desks that need daily coverage, in priority order."""

    EAST = "East"
    WEST = "West"
    VAB = "VAB"
    RESOURCE = "Resource"

    @classmethod
    def priority_order(cls) -> list["Desk"]:
        """Fill order when there aren't enough people for every desk."""
        return [cls.EAST, cls.WEST, cls.VAB, cls.RESOURCE]


@dataclass(frozen=True)
class Employee:
    """A team member who can be scheduled."""

    name: str

    def __post_init__(self) -> None:
        if not self.name or not self.name.strip():
            raise ValueError("Employee name cannot be blank")


@dataclass(frozen=True)
class ShiftRequest:
    """
    One parsed cell from the Kronos "Location Schedule" export: a person
    working a given day on a given shift (e.g. "D12", "08c12"). Kronos has
    no concept of desk assignment -- that's what this whole tool adds --
    so preferred_desk will normally be None for real data. It's kept as an
    optional hint (rather than removed) so a manually-curated request list
    could still express one, without requiring a second code path.
    """

    employee: Employee
    work_date: date
    preferred_desk: Desk | None = None
    shift_code: str = ""  # e.g. "D12", "08c12" -- informational, shown on the calendar
    notes: str = ""


@dataclass
class DeskAssignment:
    """The final result: who is on which desk, on which day."""

    work_date: date
    desk: Desk
    employee: Employee | None  # None means the desk went unfilled that day
    shift_code: str = ""  # carried over from the winning ShiftRequest, for display


@dataclass
class ShortfallFlag:
    """Raised when a desk could not be covered on a given day."""

    work_date: date
    desk: Desk
    reason: str


@dataclass
class SchedulePeriod:
    """A resolved schedule for a date range: assignments + any shortfalls."""

    start: date
    end: date
    assignments: list[DeskAssignment] = field(default_factory=list)
    shortfalls: list[ShortfallFlag] = field(default_factory=list)

    def assignments_for_day(self, day: date) -> list[DeskAssignment]:
        return [a for a in self.assignments if a.work_date == day]
