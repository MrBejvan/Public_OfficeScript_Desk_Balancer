from datetime import date

import pytest

from core.models import Desk, Employee, ShiftRequest
from core.scheduler import build_schedule


def _req(name, d, desk=None):
    return ShiftRequest(employee=Employee(name=name), work_date=d, preferred_desk=desk)


def test_fills_desks_in_priority_order_when_short_staffed():
    day = date(2026, 7, 1)
    requests = [_req("Alice", day), _req("Bob", day)]  # no preference, only 2 people
    period = build_schedule(requests, day, day)

    assigned_desks = {
        a.desk for a in period.assignments_for_day(day) if a.employee is not None
    }
    assert assigned_desks == {Desk.EAST, Desk.WEST}

    shortfall_desks = {f.desk for f in period.shortfalls}
    assert shortfall_desks == {Desk.VAB, Desk.RESOURCE}


def test_never_assigns_more_than_one_desk_per_person_per_day():
    day = date(2026, 7, 1)
    requests = [_req("Alice", day)]
    period = build_schedule(requests, day, day)
    alice_assignments = [
        a
        for a in period.assignments_for_day(day)
        if a.employee is not None and a.employee.name == "Alice"
    ]
    assert len(alice_assignments) == 1


def test_respects_explicit_desk_preference():
    day = date(2026, 7, 1)
    requests = [
        _req("Alice", day, Desk.RESOURCE),
        _req("Bob", day, Desk.EAST),
    ]
    period = build_schedule(requests, day, day)
    by_desk = {a.desk: a.employee for a in period.assignments_for_day(day)}
    assert by_desk[Desk.EAST].name == "Bob"
    assert by_desk[Desk.RESOURCE].name == "Alice"
    # Alice explicitly wanted Resource, so she should NOT be poached to
    # cover West or VAB even though they're unfilled and higher priority.
    assert by_desk[Desk.WEST] is None
    assert by_desk[Desk.VAB] is None


def test_never_exceeds_four_people_per_day():
    day = date(2026, 7, 1)
    requests = [_req(name, day) for name in ["A", "B", "C", "D", "E", "F"]]
    period = build_schedule(requests, day, day)
    filled = [a for a in period.assignments_for_day(day) if a.employee is not None]
    assert len(filled) == 4
    assert len({a.employee.name for a in filled}) == 4


def test_fairness_rotates_among_flexible_people_across_the_week():
    # Five people with no preference, every day for a week -- desks need 4
    # bodies/day, so across the week everyone should get a reasonably even
    # number of shifts rather than the same 4 people every single day.
    start = date(2026, 7, 6)  # a Monday
    end = date(2026, 7, 12)
    names = ["A", "B", "C", "D", "E"]
    requests = [_req(name, start + __import__("datetime").timedelta(days=i))
                for i in range(7) for name in names]

    period = build_schedule(requests, start, end)
    counts = {}
    for a in period.assignments:
        if a.employee is not None:
            counts[a.employee.name] = counts.get(a.employee.name, 0) + 1

    assert set(counts) == set(names)
    assert max(counts.values()) - min(counts.values()) <= 1


def test_invalid_range_raises():
    with pytest.raises(ValueError):
        build_schedule([], date(2026, 7, 5), date(2026, 7, 1))


def test_empty_requests_flags_all_desks_every_day():
    day = date(2026, 7, 1)
    period = build_schedule([], day, day)
    assert len(period.shortfalls) == 4
    assert all(a.employee is None for a in period.assignments)
