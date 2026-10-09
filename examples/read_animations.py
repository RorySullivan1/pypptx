"""Report the transitions and animations in a deck.

pypptx reads transitions and animations but does not author them, so point this script at a
deck made in PowerPoint:

    python read_animations.py deck.pptx

For each slide it prints the transition (effect, speed or duration, how the slide advances)
and each animation effect in play order: what starts it, what kind of effect it is, which
shape (and paragraphs) it animates, and its delay and duration. Run without an argument it
reports on a new blank deck, which has neither.

Demonstrates: feature-manifest.md "Transitions & Animations" (read-only).
Writes nothing.
"""

import sys

from pptx import Presentation
from pptx.enum.animation import PP_TRANSITION_TYPE


def seconds(value):
    """Format a `timedelta` in seconds, or '-' when it is None ("indefinite" or not stated)."""
    return "-" if value is None else f"{value.total_seconds():g}s"


def describe_transition(transition):
    if transition.type == PP_TRANSITION_TYPE.NONE and transition.speed is None:
        return "none"
    parts = [transition.type.name]
    if transition.direction:
        parts.append(f"dir={transition.direction}")
    if transition.duration is not None:
        parts.append(seconds(transition.duration))
    elif transition.speed is not None:
        parts.append(transition.speed.name.lower())
    if not transition.advance_on_click:
        parts.append("no click")
    if transition.advance_after is not None:
        parts.append(f"auto after {seconds(transition.advance_after)}")
    return ", ".join(parts)


def describe_animation(animation):
    effect = animation.effect_type.name if animation.effect_type else f"#{animation.preset_id}"
    kind = animation.preset_class.name if animation.preset_class else "?"
    target = animation.shape.name if animation.shape is not None else "(not a shape)"
    if animation.paragraphs is not None:
        target += f" paragraphs {animation.paragraphs.start}-{animation.paragraphs.stop - 1}"
    trigger = animation.trigger.name
    if animation.trigger_shape is not None:
        trigger += f" ({animation.trigger_shape.name})"
    return (
        f"{trigger:<22} {kind:<9} {effect:<14} {target}"
        f"  delay {seconds(animation.delay)}, takes {seconds(animation.duration)}"
    )


prs = Presentation(sys.argv[1]) if len(sys.argv) > 1 else Presentation()
if len(sys.argv) == 1:
    prs.slides.add_slide(prs.slide_layouts[0])

for number, slide in enumerate(prs.slides, start=1):
    print(f"slide {number}: transition {describe_transition(slide.transition)}")
    for animation in slide.animations:
        print(f"    {describe_animation(animation)}")

handout_master = prs.handout_master
print(
    "handout master: "
    + ("none" if handout_master is None else f"{len(handout_master.placeholders)} placeholders")
)
