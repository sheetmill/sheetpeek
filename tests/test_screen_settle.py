"""Settling a window capture (``capture.screen.window.settle``) on a scripted screen and clock."""

import pytest
from capture.screen.window import GIVE_UP, settle, settled_text

SHOT = 0.2
"""Seconds a shot takes (``screencapture -l`` measured about 0.19 s in the macOS VM)."""


class Screen:
    """Shots from a script: ``frames`` is a list of (seconds from the start, picture) changes."""

    def __init__(self, frames, shot=SHOT):
        self.frames = frames
        self.cost = shot
        self.t = 0.0
        self.taken = 0

    def clock(self):
        return self.t

    def sleep(self, seconds):
        self.t += seconds

    def shoot(self):
        picture = [p for at, p in self.frames if at <= self.t][-1]
        self.t += self.cost
        self.taken += 1
        return picture


def run(frames, shot=SHOT, **options):
    screen = Screen(frames, shot)
    shot, found = settle(screen.shoot, clock=screen.clock, sleep=screen.sleep, **options)
    return shot, found, screen


def test_a_still_window_settles_after_four_tenths_of_a_second():
    shot, found, _ = run([(0, b"a")])
    assert shot == b"a"
    assert found["shots"] == 3  # at 0, 0.2 and 0.4 s
    assert found["last_change"] == 0
    assert found["still"] >= 0.4
    assert found["interrupted"] == 0
    assert found["near"] is False


def test_a_change_starts_the_stillness_again_and_the_last_picture_is_kept():
    # "b" appears at 0.3 s: the shot started at 0.4 s sees it, the ones at 0.6 and 0.8 hold it
    shot, found, _ = run([(0, b"a"), (0.3, b"b")])
    assert shot == b"b"
    assert found["shots"] == 5
    assert found["last_change"] == pytest.approx(0.4)
    assert found["still"] == pytest.approx(0.4)
    assert found["interrupted"] == pytest.approx(0.2)  # "a" held from 0 to the shot at 0.2


def test_a_stillness_interrupted_near_the_limit_is_flagged():
    # Shots every 0.1 s: "a" holds through the shot at 0.3 s and "b" is seen at 0.4, a stillness
    # of 0.3 that a change interrupted, within 0.1 of the 0.4 that settles: flagged.
    shot, found, _ = run([(0, b"a"), (0.35, b"b")], shot=0.1)
    assert shot == b"b"
    assert found["interrupted"] == pytest.approx(0.3)
    assert found["near"] is True
    assert "near=true" in settled_text(found)


def test_a_window_that_never_holds_still_ends_the_run():
    frames = [(k * 0.2, bytes([k % 2])) for k in range(100)]
    with pytest.raises(RuntimeError, match=f"within {GIVE_UP:g} s"):
        run(frames)


def test_the_first_capture_is_shot_again_a_second_later():
    _, found, screen = run([(0, b"a")], confirm=True)
    assert found["confirmed"] is True
    assert screen.taken == 4
    with pytest.raises(RuntimeError, match="first capture"):
        run([(0, b"a"), (1.2, b"b")], confirm=True)


def test_an_audit_watches_two_seconds_and_refuses_a_change():
    _, found, screen = run([(0, b"a")], audit=True)
    assert found["audited"] is True
    assert screen.t >= 0.6 + 2.0
    with pytest.raises(RuntimeError, match="audit"):
        run([(0, b"a"), (2.0, b"b")], audit=True)
