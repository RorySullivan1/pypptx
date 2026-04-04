"""3D view and surface-related oxml objects."""

from __future__ import annotations

from pptx.oxml.simpletypes import XsdBoolean, XsdInt, XsdUnsignedInt
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class CT_View3D(BaseOxmlElement):
    """`c:view3D` element, specifying 3D view properties for a chart."""

    _tag_seq = (
        "c:rotX",
        "c:hPercent",
        "c:rotY",
        "c:depthPercent",
        "c:rAngAx",
        "c:perspective",
        "c:extLst",
    )
    rotX = ZeroOrOne("c:rotX", successors=_tag_seq[1:])
    hPercent = ZeroOrOne("c:hPercent", successors=_tag_seq[2:])
    rotY = ZeroOrOne("c:rotY", successors=_tag_seq[3:])
    depthPercent = ZeroOrOne("c:depthPercent", successors=_tag_seq[4:])
    rAngAx = ZeroOrOne("c:rAngAx", successors=_tag_seq[5:])
    perspective = ZeroOrOne("c:perspective", successors=_tag_seq[6:])
    del _tag_seq

    @property
    def rot_x_val(self):
        """Integer rotation about the X axis in degrees, or None."""
        rotX = self.rotX
        if rotX is None:
            return None
        return rotX.val

    @rot_x_val.setter
    def rot_x_val(self, value):
        self._remove_rotX()
        if value is not None:
            self._add_rotX(val=value)

    @property
    def rot_y_val(self):
        """Integer rotation about the Y axis in degrees, or None."""
        rotY = self.rotY
        if rotY is None:
            return None
        return rotY.val

    @rot_y_val.setter
    def rot_y_val(self, value):
        self._remove_rotY()
        if value is not None:
            self._add_rotY(val=value)

    @property
    def r_ang_ax_val(self):
        """Boolean right-angle axes setting, or None."""
        rAngAx = self.rAngAx
        if rAngAx is None:
            return None
        return rAngAx.val

    @r_ang_ax_val.setter
    def r_ang_ax_val(self, value):
        self._remove_rAngAx()
        if value is not None:
            self._add_rAngAx(val=value)

    @property
    def perspective_val(self):
        """Integer perspective angle (0-240), or None."""
        perspective = self.perspective
        if perspective is None:
            return None
        return perspective.val

    @perspective_val.setter
    def perspective_val(self, value):
        self._remove_perspective()
        if value is not None:
            self._add_perspective(val=value)

    @property
    def depth_percent_val(self):
        """Integer depth percent (20-2000), or None."""
        depthPercent = self.depthPercent
        if depthPercent is None:
            return None
        return depthPercent.val

    @depth_percent_val.setter
    def depth_percent_val(self, value):
        self._remove_depthPercent()
        if value is not None:
            self._add_depthPercent(val=value)

    @property
    def h_percent_val(self):
        """Integer height percent (5-500), or None."""
        hPercent = self.hPercent
        if hPercent is None:
            return None
        return hPercent.val

    @h_percent_val.setter
    def h_percent_val(self, value):
        self._remove_hPercent()
        if value is not None:
            self._add_hPercent(val=value)


class CT_RotX(BaseOxmlElement):
    """`c:rotX` element — rotation about the X axis (-90 to 90 degrees)."""

    val = OptionalAttribute("val", XsdInt, default=0)


class CT_RotY(BaseOxmlElement):
    """`c:rotY` element — rotation about the Y axis (0 to 360 degrees)."""

    val = OptionalAttribute("val", XsdUnsignedInt, default=0)


class CT_Perspective(BaseOxmlElement):
    """`c:perspective` element — perspective angle (0 to 240)."""

    val = OptionalAttribute("val", XsdUnsignedInt, default=30)


class CT_DepthPercent(BaseOxmlElement):
    """`c:depthPercent` element — depth as percentage of width (20-2000)."""

    val = OptionalAttribute("val", XsdUnsignedInt, default=100)


class CT_HPercent(BaseOxmlElement):
    """`c:hPercent` element — height as percentage of width (5-500)."""

    val = OptionalAttribute("val", XsdUnsignedInt, default=100)


class CT_Surface(BaseOxmlElement):
    """`c:floor`, `c:sideWall`, `c:backWall` elements.

    Specifies the visual properties of a chart surface (floor or wall).
    """

    _tag_seq = ("c:thickness", "c:spPr", "c:pictureOptions", "c:extLst")
    thickness = ZeroOrOne("c:thickness", successors=_tag_seq[1:])
    spPr = ZeroOrOne("c:spPr", successors=_tag_seq[2:])
    del _tag_seq


class CT_Thickness(BaseOxmlElement):
    """`c:thickness` element — thickness of a surface."""

    val = OptionalAttribute("val", XsdUnsignedInt, default=0)
