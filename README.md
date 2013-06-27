J-Interop Typesafe Proxy
========================
This small library allows you to access `IDispatch` COM objects through type-safe Java interfaces by using j-interop underneath.

Usage
=====
First, you define an interface that designates methods and properties available on the COM objects you want to talk to. Looking at the type libraries of those COM objects will help you determine the signatures.
For example,

    public interface SWbemLocator extends JIProxy {
        SWbemServices ConnectServer(String server, String namespace, String user, String password) throws JIException;
        SWbemServices ConnectServer(String server, String namespace, String user, String password, String locale, String authority, int securityFlags, Object objwbemNamedValueSet) throws JIException;
    }

Your interface must extend `JIProxy` to indicate that it is a type-safe proxy to `IDispatch`.
Methods should also throw `JIException` which comes from j-interop. The method names must match those defined by COM objects (hence the non-camel convention), and method parameters have to match the parameters defined by COM objects, too.

If the `IDispatch` interface defines properties, you need to use `@Property` annotation to mark them.

    public interface SWbemObjectSet extends JIProxy, Iterable<SWbemObject> {
        @Property
        int Count() throws JIException;
        IJIDispatch Item(String path/*?*/) throws JIException;
    }

The above example also shows the use of `Iterable<T>` in the proxy interface, which is converted to the calls to `IEnumVARIANT` COM interface.

Given those type-safe interfaces, you can wrap an `IJIDispatch` or `IJIComObject` into this proxy as follows:

    IJIComObject o = ...;
    SWbemLocator loc = JInteropInvocationHandler.wrap(SWbemLocator.class,o)

