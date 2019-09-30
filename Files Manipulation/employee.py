# -*- coding: utf-8 -*-
"""
Created on Tue Mar 19 22:39:34 2019

@author: Elvis Ma
"""

from classObject import Account

class Employee(object):
    
    """
    represent an employee
    
    Argument:
        name(string)
    attributes:
        name(string)
        job(string)
        account(Account)
    """
    def __init__(self,name):
        self.name=name
        self.job=''
        self.account=Account(name)
        
        
        
        
        
        
        