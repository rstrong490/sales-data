#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr 29 06:41:44 2020

@author: Riley Strong
"""

import pandas as pd
xl = pd.ExcelFile("OSSalesData.xlsx") 
lines = "="*60
SalesData = xl.parse("Orders")

def RegSegCount(): #1) informs you of the different segment totals for user selected region 
    print("\nEnter (1) to view East region" +
          "\nEnter (2) to view West region" +
          "\nEnter (3) to view Central region" +
          "\nEnter (4) to view South region")
    
    def RegSegCountNest(region): #This function informs you of each Segment total pertaining to each region, viewable by years
        regions = SalesData.Region.unique()
        print("This is the total of the segments pertaining to the user selectable regions:\n")
        print(regions)
        SegData = SalesData[["Segment", "Region"]]  
        RegSegData = SegData.loc[SegData["Region"]==region]
        SegDataCount = RegSegData.groupby(by="Segment").count()
        SegDataCount.rename(columns = {"Region":"Account Total"}, inplace = True)
        print(lines)
        print(region)
        print(SegDataCount)
    
    region_choice = input("Please enter a number from 1-4: ")
    if region_choice =='1':
        RegSegCountNest("East")
    elif region_choice=='2':
        RegSegCountNest("West")
    elif region_choice =='3':
        RegSegCountNest("Central")
    elif region_choice=='4':
        RegSegCountNest("South")
    else:
        print("You selected a wrong input. Please try again.")
        print(lines)
        RegSegCount()
    
def BundleOffer(): #2) informs you of products that can be bundled together to meet cross-selling goals
    print("\nProducts which can be used in bundle offers are:"+
          "\n(1)Tables"+
          "\n(2)Bookcases"+
          "\n(3)Fasteners"+
          "\n(4)Machines"+
          "\n(5)Art"+
          "\n(6)Envelopes"+
          "\n(7)Storage"+
          "\n(8)Chairs"+
          "\n(9)Binders"+
          "\n(10)Paper"+
          "\n(11)Copiers")
    bundle_option = input("\nEnter the number corresponding a specific product to see the potential bundles: ")
    if bundle_option=='1':
        print("\nPotential bundle found for Tables is"+"\nChairs")
    elif bundle_option=='2':
        print("\nPotential bundle found for Bookcases is"+"\nStorage")
    elif bundle_option=='3':
        print("\nPotential bundle found for Fasteners is"+"\nBinders")
    elif bundle_option=='4':
        print("\nPotential bundle found for Machines is"+"\nPaper")
    elif bundle_option=='5':
        print("\nPotential bundle found for Art is"+"\nPaper")
    elif bundle_option=='6':
        print("\nPotential bundle found for Envelopes is"+"\nPaper")
    elif bundle_option=='7':
        print("\nPotential bundle found for Storage is"+"\nBookcases")
    elif bundle_option=='8':
        print("\nPotential bundle found for Chairs is"+"\nTables")
    elif bundle_option=='9':
        print("\nPotential bundles found for Binders are"+"\nPaper"+"\nFasteners")
    elif bundle_option=='10':
        print("\nPotential bundles found for Paper are"+"\nBinders"
              +"\nEnvelopes"+"\nMachines"+"\nCopiers"+"\nArt")
    elif bundle_option=='11':
        print("\nPotential bundle found for Copiers is"+"\nPaper")
    else:
        print("You entered a wrong selection. Please try again.")
        print(lines)
        BundleOffer()
    
def RegSegProdFreq(): #3 informs you of the most profitable products in each segment, viewable by regions
    print("\nEnter (1) to view East" +
          "\nEnter (2) to view West" +
          "\nEnter (3) to view Central" +
          "\nEnter (4) to view South")

    def RegSegProdCount(region): #this function displays profitable products of segments in each region, using a region parameter
        RegSegRegions= SalesData[["Region", "Segment", "Product Name"]]
        RegionProdData = RegSegRegions.loc[RegSegRegions["Region"]==region]
        segments = SalesData.Segment.unique()
        
        for segment in segments:
            RegSegData = RegionProdData.loc[RegionProdData["Segment"]==segment]
            RegSegView = RegSegData[["Segment", "Product Name"]]
            RegSegProdProfit = RegSegView.groupby(by="Product Name").count().sort_values("Segment", ascending = False)
            RegSegProdProfit.rename(columns = {"Segment": "Total Times Bought"}, inplace = True)
            print(lines)
            print(segment +" "+ str(region))
            print(RegSegProdProfit.head(10))

    choice = input("Please enter your selection here: ")
    print(lines)
    if choice=='1':
        RegSegProdCount("East")
    elif choice=='2':
        RegSegProdCount("West")
    elif choice=='3':
        RegSegProdCount("Central")
    elif choice=='4':
        RegSegProdCount("South")
    else:
        print("\nYou selected a wrong input. Please try again.")
        print(lines)
        RegSegProdFreq()

def CusFreq(): #4) informs you of the 10 customers with the highest frequency of purchases in user selected year or overall.
    print("\nEnter (1) to view 2019 customer frequency" + 
          "\nEnter (2) to view 2018 customer frequency" +  
          "\nEnter (3) to view 2017 customer frequency" + 
          "\nEnter (4) to view 2016 customer frequency" +
          "\nEnter (5) to view overall customer purchase frequency")
    
    def CusFreqNest(year): #this function displays 10 customers w/ highest frequency of purchase by taking a parameter for desired year
        SalesDataYear = SalesData 
        SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
        years = SalesDataYear.Year.unique()
        print("This is the frequency of purchase of the top 10 customers pertaining to the user selectable years:\n")
        print(years)
        CusFreqData = SalesDataYear[["Customer Name", "Year", "Customer ID"]]
        CusFreqDataByYear = CusFreqData.loc[CusFreqData["Year"]==year]
        CusFreqDataTraits = CusFreqDataByYear[["Year", "Customer Name", "Customer ID"]]
        CusFreqCount = CusFreqDataTraits.groupby(by=["Customer ID", "Customer Name"]).count().sort_values("Year", ascending = False)
        CusFreqCount.rename(columns = {"Year":"Purchase Frequency"}, inplace = True)
        print(lines)
        print(year)
        print(CusFreqCount.head(10))
        
    def CusFreqOverall(): #this function informs you of overall highest customer purchase frequency
        SalesDataYear = SalesData 
        print("This is the frequency of purchase of the top 10 customers:\n")
        CusFreqData = SalesDataYear[["Customer Name", "Customer ID", "Order Date"]]
        CusFreqDataTraits = CusFreqData[["Customer Name", "Customer ID", "Order Date"]]
        CusFreqCount = CusFreqDataTraits.groupby(by=["Customer ID", "Customer Name"]).count().sort_values(by="Order Date", ascending=False)
        CusFreqCount.rename(columns = {"Order Date":"Overall Purchase Frequency"}, inplace = True)
        print(lines)
        print(CusFreqCount.head(10))
    
    year_choice = input("Please enter your selection here: ")
    if year_choice=='1':
        CusFreqNest(2019)
    elif year_choice=='2':
        CusFreqNest(2018)
    elif year_choice=='3':
        CusFreqNest(2017)
    elif year_choice=='4':
        CusFreqNest(2016)
    elif year_choice=='5':
        CusFreqOverall()
    else:
        print("You entered a wrong selection. Please try again.")
        print(lines)
        CusFreq()
    
def ShipMode(): #5 Informs you of the preferred mode of shipment in desired region
    print("\nEnter (1) to view East region" + 
          "\nEnter (2) to view West region" + 
          "\nEnter (3) to view South region" +
          "\nEnter (4) to view Central region")
    
    def ShipModeNest(region): #This function displays preferred mode of shipment by taking a parameter for desired region
        regions = SalesData.Region.unique()
        print("This is the total number of customers that opt for the different modes of shipment in the user selected regions:\n")
        print(regions)
        ShpData = SalesData[["Ship Mode", "Region"]]      
        RegShpData = ShpData.loc[ShpData["Region"]==region]
        ShpDataCount = RegShpData.groupby(by="Ship Mode").count().sort_values("Region", ascending = False)
        ShpDataCount.rename(columns = {"Region":"Count"}, inplace = True)
        print(lines)
        print(region)
        print(ShpDataCount)
        print("\nAs we can see, the preferred mode of shipment in this region is: "
          +str(max('Standard Class','Second Class','First Class','Same Day')))
    
    region_choice = input("Please enter your selection here: ")
    if region_choice=='1':
        ShipModeNest("East")
    elif region_choice=='2':
        ShipModeNest("West")
    elif region_choice=='3':
        ShipModeNest("South")
    elif region_choice=='4':
        ShipModeNest("Central")
    else:
        print("You selected a wrong input. Please try again.")
        print(lines)
        ShipMode()
    
def LoyaltyPoint(): #6) Computates monetary value of transaction to determine points earned for the loyalty program
    print("\nEvery $100 of purchase counts as 10 loyalty points in the loyalty program.")
    while True:
        try:
            transaction_amount= int(input("\nPlease enter the value of the final monetary transaction: $"))
            break;
        except:
            print("That's not a valid option!")
            print(lines)
   
    if transaction_amount<100 and transaction_amount >= 0:
        print("\nSorry! This transaction is not eligible for any loyalty points.")
    elif transaction_amount>=100:
        customer_point= int(transaction_amount/10)
        print("\nThe customer has earned "+str(customer_point)+" loyalty points from this monetary transaction.")
        print("\nCustomers will earn free shipping on their next 2 orders for every 50 loyalty points earned.")
        print(lines)
    else:
        print("\nCannot input negative numbers.")

def CityState(): #7 informs you of the top 5 highest & lowest cities & states in purchase frequency & profitability
    print("\nEnter (1) to view the top 5 purchase frequency of states & cities in 2019" +
          "\nEnter (2) to view the top 5 purchase frequency of states & cities in 2018" +
          "\nEnter (3) to view the top 5 purchase frequency of states & cities in 2017" +
          "\nEnter (4) to view the top 5 purchase frequency of states & cities in 2016" +
          "\nEnter (5) to view the top 5 lowest overall purchase frequency of states & cities" +
          "\nEnter (6) to view the top 5 most profitable states & cities in 2019" +
          "\nEnter (7) to view the top 5 most profitable states & cities in 2018" +
          "\nEnter (8) to view the top 5 most profitable states & cities in 2017" +
          "\nEnter (9) to view the top 5 most profitable states & cities in 2016" +
          "\nEnter (10) to view the top 5 least profitable states & cities in 2019" + 
          "\nEnter (11) to view the top 5 least profitable states & cities in 2018" + 
          "\nEnter (12) to view the top 5 least profitable states & cities in 2017" + 
          "\nEnter (13) to view the top 5 least profitable states & cities in 2016")
    
    def StateTopFreq(year): #This function displays top 5 states pertaining to their purchase frequency in desired year
        SalesDataYear = SalesData 
        SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
        TopStateData = SalesDataYear[["Year", "State"]]
        StateTopData = TopStateData.loc[TopStateData["Year"]==year]
        TopStateFreq = StateTopData.groupby(by = "State").count().sort_values(by="Year", ascending = False)
        TopStateFreq.rename(columns = {"Year": "Purchase Frequency"}, inplace = True)
        print(year)
        print(TopStateFreq.head(5))

    def CityTopFreq(year): #This function displays top 5 cities pertaining to their purchase frequency in desired year
        SalesDataYear = SalesData 
        SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
        TopCityData = SalesDataYear[["Year", "City"]]
        CityTopData = TopCityData.loc[TopCityData["Year"]==year]
        TopCityFreq = CityTopData.groupby(by = "City").count().sort_values(by="Year", ascending = False)
        TopCityFreq.rename(columns = {"Year": "Purchase Frequency"}, inplace = True)
        print(lines)
        print(TopCityFreq.head(5))
        
    def LowOverallStateFreq(): #This function displays states with lowest overall purchase frequency
        SalesDataYear = SalesData
        OverallStateData = SalesDataYear[["Order Date", "State"]]
        StateFreq = OverallStateData.groupby(by="State").count().sort_values(by="Order Date", ascending = True)
        StateFreq.rename(columns = {"Order Date": "Overall Lowest Purchase Frequency of States"}, inplace = True)
        print(lines)
        print(StateFreq.head(5))
        
    def LowOverallCityFreq(): #This function displays cities with lowest overall purchase frequency
        SalesDataYear = SalesData
        OverallStateData = SalesDataYear[["Order Date", "City"]]
        StateFreq = OverallStateData.groupby(by="City").count().sort_values(by="Order Date", ascending = True)
        StateFreq.rename(columns = {"Order Date": "Overall Lowest Purchase Frequency of Cities"}, inplace = True)
        print(lines)
        print(StateFreq.head(5))
    
    def StateTopProfit(year): #This function displays top 5 states pertaining to their profitability in desired year
        SalesDataYear = SalesData
        SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
        TopStateData = SalesDataYear[["Year", "State", "Profit"]]
        StateTopData = TopStateData.loc[TopStateData["Year"]==year]
        StateTopDataTraits = StateTopData[["State", "Profit"]]
        TopStateProfit = StateTopDataTraits.groupby(by = "State").sum().round(2).sort_values("Profit", ascending = False)
        print(year)
        print(TopStateProfit.head(5))
        
    def CityTopProfit(year): #This function displays top 5 cities pertaining to their profitability in desired year
        SalesDataYear = SalesData
        SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
        TopCityData = SalesDataYear[["Year", "City", "Profit"]]
        CityTopData = TopCityData.loc[TopCityData["Year"]==year]
        CityTopDataTraits = CityTopData[["City", "Profit"]]
        TopCityProfit = CityTopDataTraits.groupby(by = "City").sum().round(2).sort_values("Profit", ascending = False)
        print(lines)
        print(TopCityProfit.head(5))
        
    def StateBotProfit(year): #This function displays the top 5 most unprofitabile states in the desired year
        SalesDataYear = SalesData
        SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
        BotStateData = SalesDataYear[["Year", "State", "Profit"]]
        StateBotData = BotStateData.loc[BotStateData["Year"]==year]
        StateBotDataTraits = StateBotData[["State", "Profit"]]
        BotStateProfit = StateBotDataTraits.groupby(by = "State").sum().round(2).sort_values("Profit", ascending = True)
        print(year)
        print(BotStateProfit.head(5))
        
    def CityBotProfit(year): #This function displays the top 5 most unprofitable cities in the desired year
        SalesDataYear = SalesData
        SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
        BotCityData = SalesDataYear[["Year", "City", "Profit"]]
        CityBotData = BotCityData.loc[BotCityData["Year"]==year]
        CityBotDataTraits = CityBotData[["City", "Profit"]]
        BotCityProfit = CityBotDataTraits.groupby(by = "City").sum().round(2).sort_values("Profit", ascending = True)
        print(lines)
        print(BotCityProfit.head(5))
        
    choice = input("Please enter your selection here: ")
    print(lines)
    if choice=='1':
        StateTopFreq(2019)
        CityTopFreq(2019)
    elif choice=='2':
        StateTopFreq(2018)
        CityTopFreq(2018)
    elif choice=='3':
        StateTopFreq(2017)
        CityTopFreq(2017)
    elif choice=='4':
        StateTopFreq(2016)
        CityTopFreq(2016)
    elif choice=='5':
        LowOverallStateFreq()
        LowOverallCityFreq()
    elif choice=='6':
        StateTopProfit(2019)
        CityTopProfit(2019)
    elif choice=='7':
        StateTopProfit(2018)
        CityTopProfit(2018)
    elif choice=='8':
        StateTopProfit(2017)
        CityTopProfit(2017)
    elif choice=='9':
        StateTopProfit(2016)
        CityTopProfit(2016)
    elif choice=='10':
        StateBotProfit(2019)
        CityBotProfit(2019)
    elif choice=='11':
        StateBotProfit(2018)
        CityBotProfit(2018)
    elif choice=='12':
        StateBotProfit(2017)
        CityBotProfit(2017)
    elif choice =='13':
        StateBotProfit(2016)
        CityBotProfit(2016)
    else:
        print("\nYou selected a wrong input. Please try again.")
        print(lines)
        CityState()
        
def StateProfSubCat(): #8) Informs you of most & least profitable sub-cats in States with highest & lowest purchase frequency
    print("\nStates with highest purchase frequency: ['CA', 'NY', 'TX', 'PHI', 'IL']" +
          "\nStates with lowest purchase frequency: ['WY', 'WV', 'ND', 'ME', 'DC']" +
          "\nEnter (1) to view most & least profitable sub-categories in California" +
          "\nEnter (2) to view most & least profitable sub-categories in New York" +
          "\nEnter (3) to view most & least profitable sub-categories in Texas" +
          "\nEnter (4) to view most & least profitable sub-categories in Philadelphia" +
          "\nEnter (5) to view most & least profitable sub-categories in Illinois" +
          "\nEnter (6) to view most & least profitable sub-categories in Wyoming" +
          "\nEnter (7) to view most & least profitable sub-categories in West Virginia" +
          "\nEnter (8) to view most & least profitable sub-categories in North Dakota" +
          "\nEnter (9) to view most & least profitable sub-categories in Maine" +
          "\nEnter (10) to view most & least profitable sub-categories in Washington, D.C.")
    
    def StateSubCatProf(state): #informs you of the profitable sub-cats in the 5 states with the highest purchase frequency
        StateData = SalesData[["Sub-Category", "Profit", "State", "Category"]]
        StateInfo = StateData.loc[StateData["State"]==state]
        StateProfit = StateInfo.groupby(by="Sub-Category").sum().round(2).sort_values(by="Profit", ascending=False)
        print(state + " Most Profitable Sub-Categories:" + "\n")
        print(StateProfit.head(5))
        print(lines)
        
    def StateSubCatUnprof(state): #informs you of the unprofitable sub-cats in the 5 states with the lowest purchase frequency
        StateData = SalesData[["Sub-Category", "Profit", "State"]]
        StateInfo = StateData.loc[StateData["State"]==state]
        StateProfit = StateInfo.groupby(by="Sub-Category").sum().round(2).sort_values(by="Profit", ascending=True)
        print(state + " Least Profitable Sub-Categories:" + "\n")
        print(StateProfit.head(5))
        print(lines)

    choice = input("Please enter your selection here: ")
    print(lines)
    if choice=='1':
        StateSubCatProf("California")
        StateSubCatUnprof("California")
    elif choice=='2':
        StateSubCatProf("New York")
        StateSubCatUnprof("New York")
    elif choice=='3':
        StateSubCatProf("Texas")
        StateSubCatUnprof("Texas")
    elif choice=='4':
        StateSubCatProf("Philadelphia")
        StateSubCatUnprof("Philadelphia")
    elif choice=='5':
        StateSubCatProf("Illinois")
        StateSubCatUnprof("Illinois")
    elif choice=='6':
        StateSubCatProf("Wyoming")
        StateSubCatUnprof("Wyoming")
    elif choice=='7':
        StateSubCatProf("West Virginia")
        StateSubCatUnprof("West Virginia")
    elif choice=='8':
        StateSubCatProf("North Dakota")
        StateSubCatUnprof("North Dakota")
    elif choice=='9':
        StateSubCatProf("Maine")
        StateSubCatUnprof("Maine")
    elif choice=='10':
        StateSubCatProf("District of Columbia")
        StateSubCatUnprof("District of Columbia")
    else:
        print("You selected a wrong input. Please try again.")
        print(lines)
        StateProfSubCat()
        
def StateProd(): #9) informs you of product profitability in States with the Highest & Lowest Purchase Frequency
    print("\nStates with highest purchase frequency: ['CA', 'NY', 'TX', 'PHI', 'IL']" +
          "\nStates with lowest purchase frequency: ['WY', 'WV', 'ND', 'ME', 'DC']" +
          "\nEnter (1) to view most & least profitable products in California" +
          "\nEnter (2) to view most & least  profitable products in New York" +
          "\nEnter (3) to view most & least profitable products in Texas" +
          "\nEnter (4) to view most & least least profitable products in Philadelphia" +
          "\nEnter (5) to view most & least profitable products in Illinois" +
          "\nEnter (6) to view most & least profitable products in Wyoming" +
          "\nEnter (7) to view most & least profitable products in West Virginia" +
          "\nEnter (8) to view most & least profitable products in North Dakota" +
          "\nEnter (9) to view most & least profitable products in Maine" +
          "\nEnter (10) to view most & least profitable products in Washington, D.C.")
          
    def StateProfProd(state): #informs you of the most profitable products in states with highest & lowest purchase frequency
        StateData = SalesData[["Product Name", "Profit", "State", "Discount"]]
        StateInfo = StateData.loc[StateData["State"]==state]
        StateProfit = StateInfo.groupby(by="Product Name").sum().round(2).sort_values(by="Profit", ascending=False)
        print(state + " Most Profitable Products:" + "\n")
        print(StateProfit.head(5))
        print(lines)
        
    def StateUnprofProd(state): #informs you of the least profitable products in states with highest & lowest purchase frequency
        StateData = SalesData[["Product Name", "Profit", "State", "Discount"]]
        StateInfo = StateData.loc[StateData["State"]==state]
        StateProfit = StateInfo.groupby(by="Product Name").sum().round(2).sort_values(by="Profit", ascending=True)
        print(state + " Least Profitable Products:" + "\n")
        print(StateProfit.head(5))
        print(lines)

    choice = input("Please enter your selection here: ")
    print(lines)
    if choice=='1':
        StateProfProd("California")
        StateUnprofProd("California")
    elif choice=='2':
        StateProfProd("New York")
        StateUnprofProd("New York")
    elif choice=='3':
        StateProfProd("Texas")
        StateUnprofProd("Texas")
    elif choice=='4':
        StateProfProd("Philadelphia")
        StateUnprofProd("Philadelphia")
    elif choice=='5':
        StateProfProd("Illinois")
        StateUnprofProd("Illinois")
    elif choice=='6':
        StateProfProd("Wyoming")
        StateUnprofProd("Wyoming")
    elif choice=='7':
        StateProfProd("West Virginia")
        StateUnprofProd("West Virginia")
    elif choice =='8':
        StateProfProd("North Dakota")
        StateUnprofProd("North Dakota")
    elif choice =='9':
        StateProfProd("Maine")
        StateUnprofProd("Maine")
    elif choice=='10':
        StateProfProd("District of Columbia")
        StateUnprofProd("District of Columbia")
    else:
        print("You selected a wrong input. Please try again.")
        print(lines)
        StateProd()
        
import sys
def menu():
    print("\n" + "*"*60)
    print("\nEnter (1) to view the differentiation of accounts within regions"+
          "\nEnter (2) to view potential bundle offers available"+
          "\nEnter (3) to view profitable products in each region labeled by its segment" +
          "\nEnter (4) to view the frequency of purchase of the top 10 customers in any year or overall"+
          "\nEnter (5) to view the preferred mode of shipment in each region"+
          "\nEnter (6) to view how many loyalty points a customer earns from a certain monetary transaction"+
          "\nEnter (7) to view top 5 highest/lowest states & cities in profitability & frequency of purchase" +
          "\nEnter (8) to view sub-category profitability in the 5 states that purchase most/least frequently" +
          "\nEnter (9) to view most/least profitable products & its discount in top 5 states that purchase most/least frequently"+
          "\nEnter (10) to Exit")
    print("\n" + "*"*60)
    choice = input("Please enter your selection here: ")
    print("\n" + "*"*60)
    if choice=='1':
        RegSegCount()
        menu()
    elif choice=='2':
        BundleOffer()
        menu()
    elif choice=='3':
        RegSegProdFreq()
        menu()
    elif choice=='4':
        CusFreq()
        menu()
    elif choice=='5':
        ShipMode()
        menu()
    elif choice=='6':
        LoyaltyPoint()
        menu()
    elif choice=='7':
        CityState()
        menu()
    elif choice=='8':
        StateProfSubCat()
        menu()
    elif choice=='9':
        StateProd()
        menu()
    elif choice =='10':
        sys.exit
        print("Thank you for using the Synergy Office Solutions Sales Data Analytics System!")     
    else:
        print("\nInvalid input. Please enter an option from 1-10 below: ")
        menu()
print("Welcome to the Synergy Office Solutions Sales Data Analytics System!")
menu()