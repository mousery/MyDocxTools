from collections import defaultdict
import functools
from typing import Callable
from inspect import getcallargs, get_annotations, signature
from typeguard import check_type


_overload_registry = defaultdict(functools.partial(defaultdict, dict))

def get_overloads(func):
    """Return all defined overloads for *func* as a sequence."""
    # classmethod and staticmethod
    f = getattr(func, "__func__", func)
    if f.__module__ not in _overload_registry:
        return []
    mod_dict = _overload_registry[f.__module__]
    if f.__qualname__ not in mod_dict:
        return []
    return list(mod_dict[f.__qualname__].values())

def my_overload(func: Callable) -> Callable:
    """
    Put this decorator in front of some of the functions that
    have the same name in the same module
    but has different arg parameters,
    such that when you call this function,
    the first version of this function that fit the arg parameters will be called.
    
    Mostly copied (and heavily inspired lol) from typing.overload.
    
    
    
    For example:
    
    @my_overload
    def my_func(a: int, b: int, c: int = 0) -> int:
        return a + b + c

    @my_overload
    def my_func(a: float, b: float, c: float) -> float:
        return a * b * c
    
    @my_overload
    def my_func(a: list[int], b: list[float]) -> str:
        return str(sum([a[i] * b[i] for i in range(len(a))]))

    print(my_func(1, 2))                               # 1 + 2 + 0 = 3
    print(my_func(3.0, 4.0, 5.1))                      # 3.0 * 4.0 * 5.1 = 61.199999999999996
    print(my_func([1, 3, 5], [10.0, 15.0, 12.3]))      # str(1*10.0 + 3*15.0 + 5*12.3) = "116.5"
    print(my_func("a"))                                # raise error because fulfill none of the arg parameters of my_func.
    
    

    Args:
        func (Callable): your function

    Raises:
        TypeError: the arg doesn't fit any of the arg parameters

    Returns:
        Callable: the version of your function that fit the arg parameters
    """    
    f = getattr(func, "__func__", func)
    
    _overload_registry[f.__module__][f.__qualname__][f.__code__.co_firstlineno] = func
    
    def new_func(*args, **kwargs):
        
        overloaded_funcs: list[Callable] = get_overloads(func)
        for overloaded_func in overloaded_funcs:
            
            try:
                overloaded_callargs = getcallargs(overloaded_func, *args, **kwargs)
                overloaded_annotations = get_annotations(overloaded_func)
                if "return" in overloaded_annotations: del overloaded_annotations["return"]
                for ann, ann_type in overloaded_annotations.items():
                    check_type(overloaded_callargs[ann], ann_type)
            except:
                continue
            
            return overloaded_func(*args, **kwargs)
        
        raise TypeError(
    f"""
    The input argument ({", ".join([repr(k) for k in args] + [repr(k) + "=" + repr(v) for k, v in list(kwargs.items())])}) fulfills none of the arg parameters of function '{func.__name__}', which are:
    {"\n    ".join([" " + str(i+1) + ". " + str(signature(overloaded_func)) for i, overloaded_func in enumerate(overloaded_funcs)])}
    """
        )
    
    return new_func


if __name__ == "__main__":
    
    """
    For Demonstration:
    """
    @my_overload
    def my_func(a: int, b: int, c: int = 0) -> int:
        return a + b + c

    @my_overload
    def my_func(a: float, b: float, c: float) -> float:
        return a * b * c
    
    @my_overload
    def my_func(a: list[int], b: list[float]) -> str:
        return str(sum([a[i] * b[i] for i in range(len(a))]))


    print(my_func(1, 2))                               # 1 + 2 + 0 = 3
    print(my_func(3.0, 4.0, 5.1))                      # 3.0 * 4.0 * 5.1 = 61.199999999999996
    print(my_func([1, 3, 5], [10.0, 15.0, 12.3]))      # str(1*10.0 + 3*15.0 + 5*12.3) = "116.5"
    print(my_func("a"))                                # raise error because fulfill none of the arg parameters of my_func.
