# -*- coding: utf-8 -*-
"""
Created on Wed Mar 20 21:26:30 2019

@author: Elvis Ma
"""


def main():
    try:
        grade=float(input("Please enter your grade: "))
    except ValueError as error:
        print(str(error))
if __name__=='__main__':
        main()
        
        
        
#--------------------
    # name : iterators, generators, & iterables
        
#--------------------


def odd_generator(limit):
    current=1
    while current<limit:
       yield current
       current+=2    
       

def get_code(file):
    for line in file:
        if not line.strip().startswith("#"):
            yield line

def main():
    with open('wordstats.py','r') as program_file:
        for code_line in get_code(program_file):
            print(code_line)
            
            
            
# iterator
class FiniteOddNumbers(object):
    """
    an iterator for odd ppositive integers
    
    argument:
        limit(int)
    """
    def __init__(self,limit):
        self.current=1
        self.limit=limit
    
    def __next__(self):
        if self.current>=self.limit:
            raise StopIteration
        else:
            result=self.current
            self.current+=2
            return result
    
    def __iter__(self):
        return self
    
# iterables:

class FiniteOddIterables(object):
    """
    an iterable for odd positive integers
    
    argument:
        limit(int): upper limit on the sequence
    """
    
    def __init__(self,limit):
        self.limit=limit
    
    def __iter__(self):
        return FiniteOddNumbers(self.limit)
           
            