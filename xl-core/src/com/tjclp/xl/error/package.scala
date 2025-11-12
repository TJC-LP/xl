package com.tjclp.xl.error

/** Type alias for common result type */
type XLResult[A] = Either[XLError, A]
