package com.tjclp.xl.sheets

import com.tjclp.xl.styles.units.StyleId

/** Properties for columns */
case class ColumnProperties(
  width: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)
