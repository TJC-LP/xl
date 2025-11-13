package com.tjclp.xl.sheet

import com.tjclp.xl.style.units.StyleId

/** Properties for rows */
case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)
