"""
Balances submitted shift requests across the four desks for a given
schedule period.

Rules encoded here (confirmed with the manager this app is built for):

  * Four desks: East, West, VAB, Resource. At most one person per desk
    per day, and a desk is filled by exactly one of that day's requests.
  * Fill priority when requests are short: East and West first, then VAB,
    then Resource. If there still aren't enough people, the lower-priority
    desk is left blank and flagged rather than guessing.
  * A person who specifically requested a desk is only ever placed on that
    desk. People with no desk preference are the flexible pool used to
    fill whatever is still open, in priority order. (We deliberately do
    NOT pull someone off their requested desk to cover another desk --
    that was an explicit design choice: unfilled desks are surfaced to
    the manager instead of silently reassigning people.)
  * Fairness: when more than one eligible person is available for a desk,
    prefer whoever has worked the fewest shifts this week, then fewest
    shifts overall this period, so the workload stays balanced. The
    ~3-shifts/week target is a soft tie-breaker for now, not a hard cap --
    easy to tighten later once real usage patterns are known.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta

from core.models import (
    Desk,
    DeskAssignment,
    SchedulePeriod,
    ShiftRequest,
    ShortfallFlag,
)


def _daterange(start: date, end: date):
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def build_schedule(
    requests: list[ShiftRequest],
    start: date,
    end: date,
) -> SchedulePeriod:
    if start > end:
        raise ValueError("Schedule period start date must be on or before end date")

    by_day: dict[date, list[ShiftRequest]] = defaultdict(list)
    for req in requests:
        if start <= req.work_date <= end:
            by_day[req.work_date].append(req)

    total_assigned: dict[str, int] = defaultdict(int)
    week_assigned: dict[tuple[str, tuple[int, int]], int] = defaultdict(int)

    assignments: list[DeskAssignment] = []
    shortfalls: list[ShortfallFlag] = []

    for day in _daterange(start, end):
        day_requests = by_day.get(day, [])
        assigned_today: set[str] = set()
        iso_week = day.isocalendar()[:2]

        for desk in Desk.priority_order():
            available = [
                r for r in day_requests if r.employee.name not in assigned_today
            ]
            preferred = [r for r in available if r.preferred_desk == desk]
            flexible = [r for r in available if r.preferred_desk is None]
            pool = preferred if preferred else flexible

            if not pool:
                shortfalls.append(
                    ShortfallFlag(
                        work_date=day,
                        desk=desk,
                        reason="No one available or scheduled for this desk that day",
                    )
                )
                assignments.append(
                    DeskAssignment(work_date=day, desk=desk, employee=None)
                )
                continue

            def sort_key(r: ShiftRequest, _iso_week=iso_week):
                name = r.employee.name
                return (
                    week_assigned[(name, _iso_week)],
                    total_assigned[name],
                    day_requests.index(r),
                )

            chosen = min(pool, key=sort_key)
            name = chosen.employee.name

            assigned_today.add(name)
            total_assigned[name] += 1
            week_assigned[(name, iso_week)] += 1
            assignments.append(
                DeskAssignment(
                    work_date=day,
                    desk=desk,
                    employee=chosen.employee,
                    shift_code=chosen.shift_code,
                )
            )

    return SchedulePeriod(
        start=start, end=end, assignments=assignments, shortfalls=shortfalls
    )
