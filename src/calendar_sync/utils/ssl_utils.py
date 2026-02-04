"""SSL certificate handling utilities.

This module handles SSL certificate verification across different operating systems.
On Windows, Python's bundled certificates often don't include corporate/enterprise
CA certificates, causing SSL verification failures. The truststore package can
inject the OS's native certificate store into Python's SSL context.
"""

import logging
import platform
import sys

logger = logging.getLogger(__name__)

_ssl_initialized = False


def setup_ssl_truststore() -> bool:
    """
    Configure SSL to use the OS native certificate store.

    On Windows and macOS, this uses the truststore package to inject
    the system's certificate store into Python's SSL context. This is
    particularly useful in corporate environments with custom CA certificates.

    On Linux, the system certificates are usually already available to Python,
    but truststore can still help in some cases.

    Returns:
        True if truststore was successfully injected, False otherwise.
    """
    global _ssl_initialized

    if _ssl_initialized:
        return True

    try:
        import truststore
        truststore.inject_into_ssl()
        _ssl_initialized = True
        logger.info(f"SSL truststore injected for {platform.system()}")
        return True
    except ImportError:
        logger.warning(
            "truststore package not installed. "
            "If you encounter SSL certificate errors, install it with: pip install truststore"
        )
        return False
    except Exception as e:
        logger.warning(f"Failed to inject truststore: {e}")
        return False


def init_ssl():
    """
    Initialize SSL handling based on the current platform.

    This should be called early in application startup, before any
    HTTPS connections are made.
    """
    system = platform.system()

    if system == "Windows":
        # Windows often has issues with Python's bundled certificates
        # especially in corporate environments with custom CAs
        logger.debug("Windows detected, setting up SSL truststore...")
        setup_ssl_truststore()
    elif system == "Darwin":
        # macOS can also benefit from truststore
        logger.debug("macOS detected, setting up SSL truststore...")
        setup_ssl_truststore()
    else:
        # Linux usually works fine, but try truststore anyway if available
        logger.debug("Linux detected, attempting SSL truststore setup...")
        try:
            import truststore
            setup_ssl_truststore()
        except ImportError:
            # On Linux, this is usually fine without truststore
            logger.debug("truststore not available on Linux, using default SSL")
