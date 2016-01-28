package com.xr.export.common;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

@Target({ ElementType.TYPE, ElementType.FIELD })
public @interface MyConverter {
}
