"""

Call a function;
on failure, wait, and try the function again.
On repeated failures, wait longer between each successive attempt.
If the decorator runs out of attempts, then it gives up and raise
the previous exception or RetryError,

taken from:
https://wiki.python.org/moin/PythonDecoratorLibrary#Retry
"""

import time
import math
import logging

logger = logging.getLogger('retry')

class RetryError(Exception):
    pass

# Retry decorator with exponential backoff
def retry(tries, delay=3, backoff=2, except_retry=[]):
    """
    Retries a function or method until it returns True.

    delay sets the initial delay in seconds, and backoff sets the factor by which
    the delay should lengthen after each failure. backoff must be greater than 1,
    or else it isn't really a backoff. tries must be at least 0, and delay
    greater than 0.
    """

    assert backoff > 1, "backoff must be greater than 1"
    tries = math.floor(tries)
    assert tries >= 0, "tries must be 0 or greater"
    assert delay > 0, "delay must be greater than 0"

    def deco_retry(f):
        def f_retry(*args, **kwargs):
            mtries, mdelay = tries, delay # make mutable
            while mtries > 0:
                try:
                    result = f(*args, **kwargs)
                except Exception as e:
                    err_name = "{}.{}".format(e.__class__.__module__ ,  e.__class__.__name__)
                    err_code = None
                    err_code = [i['code'] for i in e.args if 'code' in i]
                    err_code = int(err_code[0]) if err_code else None
                    logger.debug ("exception handling {} {}".format(err_name, err_code))
                    if (err_name, err_code) in except_retry:
                        pass
                    else:
                        logger.debug ("retry.py l53: {}".format(f))
                        logger.debug ("retry.py l54: {}".format(err_name))
                        logger.debug ("retry.py l55: {}".format(err_code))
                        logger.debug ("retry.py l56: {}".format(e))
                        raise e

                    mtries -= 1      # consume an attempt
                    # logger.warning("retry {} mtries: {}".format(f, mtries))
                    if mtries > 0:
                        logger.warning("retry {} mtries: {} mdelay: {}".format(f, mtries, mdelay))
                        time.sleep(mdelay) # wait...
                        mdelay *= backoff  # make future wait longer
                        # Try again
                    else:
                        if e:
                            logger.warning ("retry.py 67: {}".format(e))
                            logger.warning ("retry.py 68: {}".format(e.args))
                            raise e
                        raise RetryError
                else:
                    return result
        return f_retry # true decorator -> decorated function
    return deco_retry  # @retry(arg[, ...]) -> true decorator

if __name__ == "__main__":

    @retry (tries=3, delay=1, backoff=2, except_retry=[('builtins.ZeroDivisionError', None)])
    def test_01():
        print ()
        for i in [1, 2, 3, 0, 4]:
            try:
                print ("{}/{}={}".format(42, i, 42/i))
            except Exception as e:
                print ("exception name: {}.{}".format(e.__class__.__module__ ,  e.__class__.__name__))
                raise e

    test_01()
