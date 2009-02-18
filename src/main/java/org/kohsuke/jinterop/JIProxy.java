package org.kohsuke.jinterop;

import org.jinterop.dcom.impls.automation.IJIDispatch;

/**
 * Base class for proxies.
 * 
 * @author Kohsuke Kawaguchi
 */
public interface JIProxy extends IJIDispatch {
    <T extends JIProxy> T cast(Class<T> type);
}
