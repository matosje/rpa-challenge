# -*- coding: utf-8 -*-
"""
Spyder Editor.

This is a temporary script file.
"""

from typing import Generator


def fib(n: int) -> Generator[int, None, None]:
    """
    Return Fibonacci sequence.

    Recieves a number and returns its repective fibonacci sequence
    """
    yield 0 # Const -- First value of Fibonacci's sequence
    yield 1 # Const -- Second value of Fibonacci's sequence

    last: int = 0
    next: int = 1
    for _ in range(1, n):
        last, next = next, last + next
        yield next


result = fib(10)
