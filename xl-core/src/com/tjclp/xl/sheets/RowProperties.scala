package com.tjclp.xl.sheets

import com.tjclp.xl.styles.units.StyleId

/** Properties for rows */
case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)
