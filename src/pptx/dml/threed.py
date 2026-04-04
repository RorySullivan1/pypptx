"""3D formatting API objects for shapes."""

from __future__ import annotations

from pptx.util import Emu, lazyproperty


class ThreeDFormat:
    """Provides access to 3D formatting properties of a shape.

    Accessed via the ``three_d`` property of a shape. Controls extrusion,
    bevels, material, and 3D scene (camera and lighting).
    """

    def __init__(self, spPr):
        self._spPr = spPr

    @property
    def bevel_bottom(self):
        """A |Bevel| object for the bottom bevel, or None if not present."""
        sp3d = self._spPr.sp3d
        if sp3d is None:
            return None
        bevelB = sp3d.bevelB
        if bevelB is None:
            return None
        return Bevel(bevelB)

    def get_or_add_bevel_bottom(self):
        """Return a |Bevel| for the bottom bevel, creating if not present."""
        sp3d = self._spPr.get_or_add_sp3d()
        return Bevel(sp3d.get_or_add_bevelB())

    @property
    def bevel_top(self):
        """A |Bevel| object for the top bevel, or None if not present."""
        sp3d = self._spPr.sp3d
        if sp3d is None:
            return None
        bevelT = sp3d.bevelT
        if bevelT is None:
            return None
        return Bevel(bevelT)

    def get_or_add_bevel_top(self):
        """Return a |Bevel| for the top bevel, creating if not present."""
        sp3d = self._spPr.get_or_add_sp3d()
        return Bevel(sp3d.get_or_add_bevelT())

    @property
    def contour_width(self):
        """Read/write |Emu| value for the contour width, or None if not set."""
        sp3d = self._spPr.sp3d
        if sp3d is None:
            return None
        val = sp3d.contourW
        if val is None:
            return None
        return Emu(val)

    @contour_width.setter
    def contour_width(self, value):
        sp3d = self._spPr.get_or_add_sp3d()
        sp3d.contourW = None if value is None else int(value)

    @property
    def extrusion_height(self):
        """Read/write |Emu| value for the extrusion depth, or None if not set."""
        sp3d = self._spPr.sp3d
        if sp3d is None:
            return None
        val = sp3d.extrusionH
        if val is None:
            return None
        return Emu(val)

    @extrusion_height.setter
    def extrusion_height(self, value):
        sp3d = self._spPr.get_or_add_sp3d()
        sp3d.extrusionH = None if value is None else int(value)

    @property
    def material(self):
        """Read/write string for the preset material, or None if not set.

        Common values: ``"warmMatte"``, ``"metal"``, ``"plastic"``,
        ``"dark"``, ``"softEdge"``, ``"flat"``, etc.
        """
        sp3d = self._spPr.sp3d
        if sp3d is None:
            return None
        return sp3d.prstMaterial

    @material.setter
    def material(self, value):
        sp3d = self._spPr.get_or_add_sp3d()
        sp3d.prstMaterial = value

    @lazyproperty
    def scene(self):
        """A |Scene3D| object for the 3D scene (camera and lighting)."""
        return Scene3D(self._spPr)


class Bevel:
    """Provides access to bevel properties (top or bottom)."""

    def __init__(self, bevel_elm):
        self._bevel = bevel_elm

    @property
    def height(self):
        """Read/write |Emu| value for the bevel height, or None if not set."""
        val = self._bevel.h
        if val is None:
            return None
        return Emu(val)

    @height.setter
    def height(self, value):
        self._bevel.h = None if value is None else int(value)

    @property
    def preset(self):
        """Read/write string for the bevel preset type, or None if not set.

        Common values: ``"circle"``, ``"relaxedInset"``, ``"angle"``,
        ``"cross"``, ``"convex"``, ``"coolSlant"``, ``"divot"``,
        ``"riblet"``, ``"hardEdge"``, ``"artDeco"``, etc.
        """
        return self._bevel.prst

    @preset.setter
    def preset(self, value):
        self._bevel.prst = value

    @property
    def width(self):
        """Read/write |Emu| value for the bevel width, or None if not set."""
        val = self._bevel.w
        if val is None:
            return None
        return Emu(val)

    @width.setter
    def width(self, value):
        self._bevel.w = None if value is None else int(value)


class Scene3D:
    """Provides access to 3D scene properties (camera and lighting)."""

    def __init__(self, spPr):
        self._spPr = spPr

    @property
    def camera(self):
        """A |Camera| object, or None if no scene3d element is present."""
        scene3d = self._spPr.scene3d
        if scene3d is None:
            return None
        camera = scene3d.camera
        if camera is None:
            return None
        return Camera(camera)

    def get_or_add_camera(self):
        """Return a |Camera| object, creating scene3d and camera if needed."""
        scene3d = self._spPr.get_or_add_scene3d()
        return Camera(scene3d.get_or_add_camera())

    @property
    def light_rig(self):
        """A |LightRig| object, or None if no scene3d element is present."""
        scene3d = self._spPr.scene3d
        if scene3d is None:
            return None
        lightRig = scene3d.lightRig
        if lightRig is None:
            return None
        return LightRig(lightRig)

    def get_or_add_light_rig(self):
        """Return a |LightRig| object, creating scene3d and lightRig if needed."""
        scene3d = self._spPr.get_or_add_scene3d()
        return LightRig(scene3d.get_or_add_lightRig())


class Camera:
    """Provides access to 3D camera properties."""

    def __init__(self, camera_elm):
        self._camera = camera_elm

    @property
    def field_of_view(self):
        """Read/write float for the field of view in degrees, or None.

        Stored internally in 60,000ths of a degree.
        """
        fov = self._camera.fov
        if fov is None:
            return None
        return fov / 60000.0

    @field_of_view.setter
    def field_of_view(self, value):
        self._camera.fov = None if value is None else int(value * 60000)

    @property
    def preset(self):
        """Read/write string for the camera preset, or None if not set.

        Common values: ``"orthographicFront"``, ``"perspectiveFront"``,
        ``"perspectiveAbove"``, ``"isometricOffAxis1Left"``, etc.
        """
        return self._camera.prst

    @preset.setter
    def preset(self, value):
        self._camera.prst = value


class LightRig:
    """Provides access to 3D lighting rig properties."""

    def __init__(self, lightRig_elm):
        self._lightRig = lightRig_elm

    @property
    def direction(self):
        """Read/write string for the light direction, or None if not set.

        Common values: ``"t"`` (top), ``"b"`` (bottom), ``"l"`` (left),
        ``"r"`` (right), ``"tl"``, ``"tr"``, ``"bl"``, ``"br"``.
        """
        return self._lightRig.dir

    @direction.setter
    def direction(self, value):
        self._lightRig.dir = value

    @property
    def rig_type(self):
        """Read/write string for the lighting rig preset, or None if not set.

        Common values: ``"threePt"``, ``"balanced"``, ``"flat"``,
        ``"soft"``, ``"harsh"``, ``"flood"``, ``"contrasting"``,
        ``"morning"``, ``"sunrise"``, ``"sunset"``, etc.
        """
        return self._lightRig.rig

    @rig_type.setter
    def rig_type(self, value):
        self._lightRig.rig = value
