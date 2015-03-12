package com.xr.export.common;

import org.springframework.beans.BeansException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Lazy;
import org.springframework.stereotype.Component;


/**
 * @author sw
 * 这个类是为了保存Spring上下文以供全局使用，
 * 但是根据原则，尽量少使用，而是通过依赖注入
 */
@Configuration
@Component
public class SpringContextUtils implements ApplicationContextAware{
	
	@Autowired
	private static ApplicationContext application;
	
	public static ApplicationContext getSpringContext() 
	{		
		if( application==null ){
			 
		}
		
		return application;
	}

	@Override
	public void setApplicationContext(ApplicationContext applicationContext)
			throws BeansException {
		this.application = applicationContext;
		
	}
}
