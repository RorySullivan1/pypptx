Units, utilities and exceptions
===============================

.. _EMU:

Length units (EMU)
------------------

Every length in the API is a |Length|: an ``int`` counting English Metric Units (EMU), the
unit Office Open XML stores lengths in. There are 914,400 EMU to the inch, 360,000 to the
centimeter and 12,700 to the point, so every common unit converts to a whole number. Pass
lengths using the subclasses below, e.g. ``Inches(1)`` or ``Pt(18)``, and read them back in
any unit with ``.inches``, ``.cm``, ``.pt`` and so on.

``pptx.util``
-------------

.. automodule:: pptx.util

``pptx.exc``
------------

.. automodule:: pptx.exc
