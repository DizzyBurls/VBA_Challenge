# VBA_Challenge

    Homework Task 2
    Writen by Chris Burley

# Name of Project:

    Share Market Analysis Tool.
    Module name: ShareMarket_CJB

# Description:

    In this project I have created a module that can, for each YEAR of share market data:

    (a) Identify the opening price (on Jan 1) of each individual share and its closing prices (on Dec 31) 
    
    See Picture_1.jpeg


    (b) Determine whether the prices of individual shares have increase or decreased between Jan 1 and Dec 31st.
    Those that have increase in price have their tickers displayed in a GREEN cell.
    Those that have increase in price have their tickers displayed in a RED cell.
    Those that have remained the same in price have their tickers displayed in a GREY cell.

    (c) Determine the percentage change in individual share prices between Jan 1 and Dec 31st.

    (d) Determine the total stock volume for an individual share between Jan 1st and Dec 31st.
    
    (e) Determine the number of days each share was traded.
    
    See Picture_2.jpeg
    
    (f) Determine the number of different share names that were traded.
    
    See Picture_3.jpeg

    
    (g) Identify the share which demonstrated the largest % increase.
        Identify the share which demonstrated the largest % decrease.
        Identify the share which had the largest share volume.
        
    See Picture_4.jpeg

    In addition to this, the module is also capable of determining:
       
    (a) The share which demonstrated the largest % increase over the ENTIRE duration that data was collected.
        (A bug exist here that I was unable to resolve in the time given. Feedback on this would be much appreciated!)
        
    (b) The share which demonstrated the largest % decrease over the ENTIRE duration that data was collected..
        
    (c) The share which had the largest share volume over the ENTIRE duration that data was collected.
    
    Again, see Picture_1.jpeg

# Installation

    To istall this module, 

    (1) Simply download the ".bas" file I have added to this GitHub repository.
    
    (2) Open the database that you wish to run the module on. (Ensure that Macros are enabled.)
    
    (3) Click on the Developer menu.
    
    (4) Click on Visual Basic
    
    (5) Right-click on the VBA project that you wish to add the module to and select "Import".
    
    (6) Choose the ".bas" file and perform the import.
    
    (7) The module you have imported should now appear on your module list for the VBA project in question.
    
# Usage

    This module is designed to be run when the active worksheet is the FIRST in the workbook.
    
    If, for example, your data set contains worksheets of stock market data from 2018 then 2019 and finally from 2020 then it is important that you click somewhere on the  2018 sheet before running the module.
    
# Support

    For any support with this module, please contact me at chrisjburley@gmail.com
    I would be more than happy to help.

# Contributions

    As this is a piece of work that is going to be assessed, i think that it would be wise for me to suggest that contributions NOT be made to this module.

# Authors and Acknowledgments

    Author - Chris Burley (CJB)

    I would like to thank Yang from the #ASkBCS for spotting an error in my "for ws in worksheets" loop.
    I would also like to acknowledge Akash for suggesting that I change some of my variables from Long to Double. I was experiencing overflow errors.

