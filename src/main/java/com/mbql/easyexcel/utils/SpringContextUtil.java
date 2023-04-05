package com.mbql.easyexcel.utils;

import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.lang.NonNull;

/**
 * Spring 上下文工具类
 *
 * @author slp
 */
public class SpringContextUtil implements ApplicationContextAware {

    private static ApplicationContext applicationContext;

    @Override
    public void setApplicationContext(@NonNull ApplicationContext applicationContext) throws BeansException {
        if (SpringContextUtil.applicationContext == null) {
            SpringContextUtil.applicationContext = applicationContext;
        }
    }

    /**
     * 获取ApplicationContext
     *
     * @return org.springframework.context.ApplicationContext
     */
    public static ApplicationContext getApplicationContext() {
        return applicationContext;
    }

    /**
     * 根据名称获取Bean对象
     *
     * @param name 名称
     * @param <T>  T
     * @return bean
     */
    public static <T> T getBean(String name) {
        //noinspection unchecked
        return (T) applicationContext.getBean(name);
    }

    /**
     * 根据class获取Bean对象
     *
     * @param clazz class
     * @param <T>   T
     * @return bean
     */
    public static <T> T getBean(Class<T> clazz) {
        return applicationContext.getBean(clazz);
    }

    /**
     * 根据名称和class获取Bean对象
     *
     * @param name  名称
     * @param clazz class
     * @param <T>   T
     * @return bean
     */
    public static <T> T getBean(String name, Class<T> clazz) {
        return applicationContext.getBean(name, clazz);
    }

    /**
     * 获取配置文件数据
     *
     * @param key key
     * @return java.lang.String
     */
    public static String getProperty(String key) {
        return applicationContext.getEnvironment().getProperty(key);
    }

    /**
     * 注册bean
     *
     * @param beanName bean名称
     * @param bean     bean
     * @param <T>      T
     */
    public static <T> void registerBean(String beanName, T bean) {
        ConfigurableApplicationContext context = (ConfigurableApplicationContext) applicationContext;
        context.getBeanFactory().registerSingleton(beanName, bean);
    }
}
