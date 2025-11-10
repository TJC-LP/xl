package com.tjclp.xl

import com.tjclp.xl.style.StyleId

/** Properties for columns */
case class ColumnProperties(
  width: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)

/** Properties for rows */
case class RowProperties(
  height: Option[Double] = None,
  hidden: Boolean = false,
  styleId: Option[StyleId] = None
)
