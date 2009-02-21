package org.kohsuke.jinterop;

import org.jinterop.dcom.impls.automation.IJIDispatch;
import org.jinterop.dcom.impls.automation.IJIEnumVariant;
import org.jinterop.dcom.impls.JIObjectFactory;
import org.jinterop.dcom.core.IJIComObject;
import org.jinterop.dcom.core.JIArray;
import org.jinterop.dcom.core.JIVariant;
import org.jinterop.dcom.core.JIString;
import org.jinterop.dcom.common.JIException;
import org.jvnet.tiger_types.Types;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.lang.reflect.Method;
import java.lang.reflect.InvocationTargetException;
import java.util.Iterator;

/**
 * @author Kohsuke Kawaguchi
 */
public class JInteropInvocationHandler implements InvocationHandler {
    public static <T extends JIProxy> T wrap(Class<T> type, IJIDispatch obj) {
        return type.cast(Proxy.newProxyInstance(type.getClassLoader(), new Class[]{type}, new JInteropInvocationHandler(obj,type)));
    }

    public static <T extends JIProxy> T wrap(Class<T> type, IJIComObject obj) throws JIException {
        return wrap(type,(IJIDispatch) JIObjectFactory.narrowObject(obj.queryInterface(IJIDispatch.IID)));
    }

    /**
     * j-interop object that actually serves the request.
     */
    private final IJIDispatch core;
    /**
     * Proxy implements this interface.
     */
    private final Class<? extends JIProxy> interfaceType;

    /**
     * If we have a delegate implementation class, that class.
     */
    private final Class staticImplementation;

    public JInteropInvocationHandler(IJIDispatch core, Class<? extends JIProxy> interfaceType) {
        this.core = core;
        this.interfaceType = interfaceType;
        this.staticImplementation = findStaticImplementation();
    }

    private Class findStaticImplementation() {
        try {
            return interfaceType.getClassLoader().loadClass(interfaceType.getName()+"$Implementation");
        } catch (ClassNotFoundException e) {
            return null;
        }
    }

    public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
//        JIVariant result = swbemLocator.callMethodA("ConnectServer",new Object[]{new JIString(address),JIVariant.OPTIONAL_PARAM(),JIVariant.OPTIONAL_PARAM(),JIVariant.OPTIONAL_PARAM()
//                ,JIVariant.OPTIONAL_PARAM(),JIVariant.OPTIONAL_PARAM(), 0,JIVariant.OPTIONAL_PARAM()})[0];

        // let the core object handle all IJIDispatch methods
        if(method.getDeclaringClass()==IJIDispatch.class) {
            return method.invoke(core,args);
        }

        if(method.getDeclaringClass()==JIProxy.class) {
            // cast method
            return wrap((Class)args[0],core);
        }

        if(method.getDeclaringClass()==Iterable.class) {
            // _NewEnum call
            IJIComObject object2 = core.get("_NewEnum").getObjectAsComObject();
            final IJIEnumVariant enumVARIANT = (IJIEnumVariant) JIObjectFactory.narrowObject(object2.queryInterface(IJIEnumVariant.IID));
            final Class expectedType = Types.erasure(Types.getTypeArgument(Types.getBaseClass(interfaceType,Iterable.class),0));
            return new Iterator() {
                Object next;
                boolean end=false;
                public boolean hasNext() {
                    fetch();
                    return next!=null;
                }

                public Object next() {
                    fetch();
                    Object r = next;
                    next = null;
                    return r;
                }

                private void fetch() {
                    try {
                        if(next!=null || end)  return;

                        Object[] values = enumVARIANT.next(1);
                        if(values.length==0)    return;
                        if(values.length!=2)
                            throw new AssertionError("Returned "+values.length);

                        Object[] ai = (Object[])((JIArray) values[0]).getArrayInstance();
                        if(ai.length!=1)
                            throw new AssertionError(ai);
                        next = unmarshal((JIVariant)ai[0],expectedType);
                    } catch (JIException e) {
                        if(e.getErrorCode()==1) {
                            end = true;
                            return;
                        }
                        throw new RuntimeException(e); // TODO: define a proper tunneling exception
                    }
                }

                public void remove() {
                    throw new UnsupportedOperationException();
                }
            };
        }

        if(method.getAnnotation(Property.class)!=null) {
            // property call
            return unmarshal(core.get(method.getName()),method.getReturnType());
        } else {
            // method call
            Class<?>[] paramTypes = method.getParameterTypes();

            if(staticImplementation!=null){// do we have a static implementation?
                Class[] staticParams = new Class[paramTypes.length+1];
                System.arraycopy(paramTypes,0,staticParams,1,paramTypes.length);
                staticParams[0] = interfaceType;

                try {
                    Method stm = staticImplementation.getMethod(method.getName(), staticParams);

                    Object[] staticArgs = new Object[staticParams.length];
                    if(args!=null)
                    System.arraycopy(args,0,staticArgs,1,args.length);
                    staticArgs[0] = proxy;

                    return stm.invoke(null,staticArgs);
                } catch (NoSuchMethodException e) {
                    // fall through
                }
            }

            // massage argument
            for (int i = 0; i < paramTypes.length; i++) {
                args[i] = marshal(args[i],paramTypes[i]);
            }

            JIVariant[] r = core.callMethodA(method.getName(), args);

            return unmarshal(r[0],method.getReturnType());
        }
    }

    private Object marshal(Object v, Class<?> declaredType) throws JIException {
        if(v==null) return JIVariant.OPTIONAL_PARAM();

        Class<?> actualType = v.getClass();
        if(actualType==String.class)
            return new JIString((String)v);
        return v;
    }

    private Object unmarshal(JIVariant v, Class<?> returnType) throws JIException {
        if(returnType==IJIDispatch.class) {
            return JIObjectFactory.narrowObject(v.getObjectAsComObject());
        }
        if(JIProxy.class.isAssignableFrom(returnType)) {// typed wrapper
            return wrap(returnType.asSubclass(JIProxy.class),v.getObjectAsComObject());
        }
        if(returnType==JIVariant.class)
            return v;
        if(returnType==int.class || returnType==Integer.class)
            return v.getObjectAsInt();
        if(returnType==void.class)
            return null;
        if(returnType==String.class)
            return v.getObjectAsString2();

        throw new UnsupportedOperationException(returnType.getName());
    }
}
