"""Custom show (`p:custShowLst`) object model.

A custom show is a named, ordered subset of the slides in a presentation. This module
provides the |CustomShows| collection and |CustomShow| item, proxies for the
`p:custShowLst` and `p:custShow` elements respectively.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, Sequence

from pptx.exc import SlideError
from pptx.shared import ParentedElementProxy

if TYPE_CHECKING:
    from pptx.oxml.presentation import CT_CustomShow, CT_Presentation
    from pptx.parts.presentation import PresentationPart
    from pptx.presentation import Presentation
    from pptx.slide import Slide


class CustomShows(ParentedElementProxy):
    """Sequence of |CustomShow| objects belonging to a |Presentation|.

    Supports `len()`, iteration, and indexed access. Custom shows are backed by the
    `p:custShowLst` element, a child of `p:presentation`.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, prs_elm: CT_Presentation, prs: Presentation):
        super(CustomShows, self).__init__(prs_elm, prs)
        self._prs_elm = prs_elm
        self._prs = prs

    def __getitem__(self, idx: int) -> CustomShow:
        """Provide indexed access, e.g. `custom_shows[0]`."""
        try:
            custShow = self._custShow_lst[idx]
        except IndexError:
            raise IndexError("custom show index out of range")
        return CustomShow(custShow, self._prs)

    def __iter__(self) -> Iterator[CustomShow]:
        """Support iteration, e.g. `for custom_show in custom_shows:`."""
        for custShow in self._custShow_lst:
            yield CustomShow(custShow, self._prs)

    def __len__(self) -> int:
        """Support len() built-in function, e.g. `len(custom_shows) == 2`."""
        return len(self._custShow_lst)

    def add(self, name: str, slides: Sequence[Slide]) -> CustomShow:
        """Return a new |CustomShow| named `name`, containing `slides` in the given order.

        `slides` must belong to this presentation. Raises |SlideError| if any of them do not.
        """
        custShowLst = self._prs_elm.get_or_add_custShowLst()
        custShow = custShowLst.add_custShow(name=name, id=self._next_id)
        sldLst = custShow.get_or_add_sldLst()
        for slide in slides:
            sldLst.add_sld(self._rId_for_slide(slide))
        return CustomShow(custShow, self._prs)

    def get(self, name: str, default: CustomShow | None = None) -> CustomShow | None:
        """Return the |CustomShow| named `name`, or `default` if no such custom show exists."""
        for custom_show in self:
            if custom_show.name == name:
                return custom_show
        return default

    def remove(self, custom_show: CustomShow) -> None:
        """Remove `custom_show` from this collection.

        Raises |SlideError| if `custom_show` is not present in this collection.
        """
        custShowLst = self._prs_elm.custShowLst
        if custShowLst is None or custom_show._element not in custShowLst.custShow_lst:
            raise SlideError("custom_show not in this collection")
        custShowLst.remove(custom_show._element)

    def _rId_for_slide(self, slide: Slide) -> str:
        """Return the rId this presentation's `p:sldIdLst` uses to refer to `slide`."""
        rId = self.part.rId_for_slide(slide.part)
        if rId is None:
            raise SlideError("slide is not in this presentation")
        return rId

    @property
    def _custShow_lst(self) -> list[CT_CustomShow]:
        custShowLst = self._prs_elm.custShowLst
        if custShowLst is None:
            return []
        return custShowLst.custShow_lst

    @property
    def _next_id(self) -> int:
        """The next available custom-show id, one greater than the current maximum in use."""
        used_ids = [custShow.id for custShow in self._custShow_lst]
        return max(used_ids, default=-1) + 1


class CustomShow(ParentedElementProxy):
    """A named, ordered subset of the slides in a presentation.

    Not intended to be constructed directly. Use `Presentation.custom_shows.add()` to
    create a new custom show.
    """

    part: PresentationPart  # pyright: ignore[reportIncompatibleMethodOverride]

    def __init__(self, custShow: CT_CustomShow, prs: Presentation):
        super(CustomShow, self).__init__(custShow, prs)
        self._custShow = custShow

    @property
    def id(self) -> int:
        """Read-only unique non-negative integer identifying this custom show."""
        return self._custShow.id

    @property
    def name(self) -> str:
        """Name of this custom show. Read/write."""
        return self._custShow.name

    @name.setter
    def name(self, value: str) -> None:
        self._custShow.name = value

    @property
    def slides(self) -> list[Slide]:
        """List of |Slide| objects in this custom show, in order."""
        sldLst = self._custShow.sldLst
        if sldLst is None:
            return []
        return [self.part.related_slide(rId) for rId in sldLst.sld_rIds()]

    @slides.setter
    def slides(self, slides: Sequence[Slide]) -> None:
        sldLst = self._custShow.get_or_add_sldLst()
        sldLst.clear_slds()
        for slide in slides:
            rId = self.part.rId_for_slide(slide.part)
            if rId is None:
                raise SlideError("slide is not in this presentation")
            sldLst.add_sld(rId)
