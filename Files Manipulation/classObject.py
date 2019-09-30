class Car(object):
    
    def __init__ (self,make, model, fuel_efficiency, mileage=0, gas=0):
        self.make=make
        self.model=model
        self.fuel_efficiency=fuel_efficiency
        self.mileage=mileage
        self.gas_in_tank=gas
    
    def __str__(self):
        return self.make+" "+self.model
    
    def tank_empty(self):
        return self.gas_in_tank==0

    def add_gas(self, amount):
        self.gas_in_tank+=amount
    
    @property
    def distance_can_go(self):
        return self.gas_in_tank*self.fuel_efficiency
    
    def drive(self,distance):
        max_distance=self.gas_in_tank*self.fuel_efficiency
        if distance <=max_distance:
            self.mileage+=distance
            self.gas_in_tank-=distance/self.fuel_efficiency
        else:
            self.mileage+=max_distance
            self.gas_in_tank=0
        return self

    
    
class Book(object):
    
    """represent a book
        argument:
            author (string)
            title(string)
        attributes:
            author(string)
            title(string)
            content(list): list containing the content of each chapter
    """
    
    def __init__(self, author, title):
        self.author=author
        self.title=title
        self.content=[]
        
    def __str__(self):
        result=self.title+"by: "+self.author
        chapter_number=1
        #add chapter numbers to the representation
        for chapter in self.content:
            result+='\nChapter'+str(chapter_number)+'\n'+chapter
            chapter_number+=1
        return result
    
    def __getitem__(self,key):
        #if the index is in the existing chapters range
        if 0<key<=len(self.content):
            return self.content[key-1] # conver to 0 based indexing
            
            
    def add_chapter(self,text):
        self.content.append(text)
            
    
    
class Account(object):
    """
    represent bank account
    argument:
        account_holder (string)
    attributes:
        holder(string)
        balance(float)
    """
    def __init__(self,account_holder):
        self.holder=account_holder
        self.balance=0
        
    def __str__(self):
        return self.holder+": $"+str(self.balance)
    
    def __add__(self,other):
        new_holder=self.holder+" & "+other.holder
        new_balance=self.balance+other.balance
        new_account= Account(new_holder)
        new_account.deposit(new_balance)
        return new_account
    
    def deposit(self, amount):
        """
        parameter: amount (float)
        """
        self.balance+=amount
        return self
    
    def withdraw(self, amount):
        if self.balance>=amount:
            self.balance=self.balance-amount
            return True
        else:
            return False
        
    
            
class SavingsAccount(Account):
      """
      represent a savings bank account with a withdrawal fee
      """    
      fee=1
      def withdraw(self,amount):
        if self.balance >= amount + self.fee:   #pay attention to the self.fee
           self.balance = self.balance - amount - self.fee
           return True
        else:
            return False

class PremiumAccount(Account):
    """
    represent a premium interest bearing bank account
    
    Argument:
        account_holder(str)
        rate(float)
    
    attributes:
        holder(str)
        balance(float)
        interest_rate(float)
    """
    def __init__(self,account_holder,rate):
        self.interest_rate=rate
        super().__init__(account_holder)   ## mehtod 1
        Account.__init__(self,account_holder) ## method 2
        
class PremiumSavingsAccount(PremiumAccount,SavingsAccount):
      """
      represent a premium interest bearing bank account with a withdrawal fee
      """        
        
        
class Account(object):
      
    def __init__(self,account_holder):
        self.holder=account_holder
        self._balance=0
    
    def deposit(self,amount):
        self._balance+=amount
        return self
    @property
    def balance(self):
        return self._balance
    
        
class Student(object):
    
    """
    represent a student
    
    argument:
        name(string)
        sid(int)  -8 digit
    
    attributes:
        name(string)
        sid(int)
    """
    number_of_students=0
    def __init__(self,name,sid):
        self.name=name
        if self.valid(sid):
            self.sid=sid
        else:
            self.sid=0
    
    @staticmethod
    def valid(some_id):
        # a valid student id starts with 2019
        if some_id//10000==2019:
            return True
        else:
            return False
        
    @classmethod
    def update_count(cls):
        cls.number_of_students +=1
        
        
        
class OddNumbers(object):
    def __init__(self):
        self.count=1
    
    def __next__(self):
        result=self.count
        self.count+=2
        return result
    
    def __iter__(self):
        return self
    def __str__(self):
        return str(self.count)
     
    # generator
    
        

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
    
    
    


    
     
        
        
        
        