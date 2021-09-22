# -*- coding: utf-8 -*-
"""
Created on Wed Sep 22 13:30:23 2021

@author: somnath
"""
import func as fun
 

while(True) :
    user_choice=int(input("which operation do you want to perform press\n 1:insert  2:check eligibity  3:update\n"))

    if user_choice==1 :
        fun.insert_data()

    elif user_choice==2:
        fun.update_data()
    
    elif user_choice==3:
        fun.write_into_sheet()

    else:
        print("entered wrong choice")
    ch=input("enter 'Y' to do task again")
    if ch!='Y' :
        break
        
    