#!/usr/bin/python

'''PuLP is an LP modeler written in python. PuLP can generate MPS or LP files
and call GLPK, COIN CLP/CBC, CPLEX, and GUROBI to solve linear problems.
https://pythonhosted.org/PuLP/'''

from pulp import *

'''win32com.client is library  that assist with interacting with Excel.'''

import win32com.client

import traceback

def test_kernels_installed():
    
    '''This function tests the availability the optimization kernels installed
    in which Pulp utilizes for large complex problems'''
    
    pulp.pulpTestAll()
    raise SystemExit(0)
#test_kernels_installed()


def get_excel(filename, ws_name):
    
    '''Given the Excel filename and the worksheet name, retrieve the Excel,
    Workbook and Worksheet objects requested. If an error occurs in retrieving
    the objects, print an error and abort the process.'''
    
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        wb_filepath = xl.ActiveWorkbook.FullName
        wb_filename = wb_filepath.split("\\")[-1]
        if wb_filename != filename:
            wb = xl.Workbooks.Open(filename)
    except Exception as e:
        traceback.print_exc()
##        print("An error occurred accessing Excel. \n" +
##              "Please ensure dialog boxes are closed and no cells are currently being edited.")
        raise SystemExit(0)
    wb = xl.ActiveWorkbook
    ws = wb.Worksheets(ws_name)
    return xl, wb, ws


def get_vaiables(wb, ws, prin = False):
    
    '''Given the worksheet object and the list of variables requested, get the
    values of each of the variables as defined by the named range in the Excel
    worksheet.'''

    '''Extract Allocation_Pref'''
    
    ws_pur = wb.Worksheets("Deals")
    x = ws_pur.Range("Allocation_Pref").Value                       #tuple
    print(x)
##    allocation_pref = [int(x[0][i]) for i in range(len(x[0]))]             #list

    allocation_pref = []
    for i in range(len(x[0])):
        if x[0][i] == "Prepay":
            ap = 0
        elif x[0][i] == "PPA":
            ap = 1
        else:
            print("Deal type error")
            raise SystemExit(0)
        allocation_pref.append(ap)   
    
    '''Extract Deals'''
    x = ws.Range("Deals").Value
    deals = [("D"+str(i), int(x[i][0])) for i in range(len(x))]

    '''Extract Deals_Type'''       
    x = ws.Range("Deals_Type").Value
    deals_type = []
    for i in range(len(x)):
        if x[i][0] == "Prepay":
            d_type = 0
        elif x[i][0] == "PPA":
            d_type = 1
        else:
            print("Deal type error") 
            raise SystemExit(0)
        deals_type.append(("D"+str(i), d_type))
    
    '''Extract Purchasers'''       
    x = ws.Range("Purchasers_Amounts").Value
    purchasers = [int(x[0][i]) for i in range(len(x[0]))]

    '''Extract Target_Gap'''       
##    target_gap = ws.Range("Target_Gap").Value
    target_gap = 0.01

    '''Extract min_deal'''
##    min_deal = ws.Range("Min_deal").Value
    min_deal = True

    '''Extract pref_penalty'''
##    pref_penalty = ws.Range("Pref_penalty").Value
    pref_penalty = True

    if prin:
        print("\nAllocation_Pref: %s" % allocation_pref)
        print("\nDeals: %s" % deals)
        print("\nDeal_Type: %s" % deals_type)
        print("\nPurchasers: %s" % purchasers)
        print("\nTarget_Gap: %s" % target_gap)
        print("\nmin_deal: %s" % min_deal)
        print("\npref_penalty: %s" % pref_penalty)
        print("\n")

    return allocation_pref, deals, deals_type, purchasers, target_gap, min_deal, pref_penalty


def put_optimum_alloaction(ws, optimum_alloaction):
    '''Given the worksheet object and the list of optimum alloactions found
    paste the values into the Excel Worksheet named range
    'Optimum_Alloaction'''
    optimum_alloaction = [[item] for idx, item in enumerate(optimum_alloaction)]
    ws.Range("Optimum_Alloaction").Value = optimum_alloaction


def format_output(k, deal_count, prin = False):
    '''This function takes a list of deal allocations to purchasers and returns a
    complete list including the deals not allocated.
    Steps:
    Convert the list to a dictionary, if the deal is in the dict append
    the set (deal ID, purchasers ID ) otherwise append (deal ID, 0) to the
    output list.'''
    k_dict = dict(k)
    m, n = [], []
    for i in range(deal_count):
        if "D%s" % i in k_dict.keys():
            m.append(("D%s" % i, k_dict["D%s" % i]+1))
            n.append(k_dict["D%s" % i]+1)
        else:
            m.append(("D%s" % i, 0))
            n.append(0)
##        if prin: print(m[i][1])
    return m, n


def optimize_deal_allocation(deals, purchasers, deals_type, allocation_pref, target_gap, \
                             min_deal, pref_penalty, time_limit, min_deal_value, prin = False):
    """
    Function to find the optimal allocation of deals to purchasers under a number
    of constraints. This function is similar to the Solve function in Excel but
    utilizes the GLPK package in order to optimise the process.

    Input:
        deals - A list of sets listing the deals identifier and size
        purchasers - A list of the purchasers max allocation amounts
        deals_type - A list of sets listing the deals identifier and the deal type 
        allocation_pref - A list of the purchasers allocation preferences between type 1 / 2
        target_gap - A float of the minimum acceptable gap between the solution found and the LP relaxation (not in use)
        min_deal - Boolean, if each purchaser is to have minimum one deal allocated to them
        pref_penalty - Boolean, if the Purchasers preferences between Type 1 and 2 should be considered
    """ 
    
    deal_count = len(deals)
    num_of_purchasers = len(purchasers)

    y = pulp.LpVariable.dicts('BinUsed', range(num_of_purchasers), lowBound = 0, upBound = 1, cat = pulp.LpInteger)
    possible_ItemInBin = [(itemTuple[0], binNum) for itemTuple in deals for binNum in range(num_of_purchasers)]
    x = pulp.LpVariable.dicts('itemInBin', possible_ItemInBin, lowBound = 0, upBound = 1, cat = pulp.LpInteger)

    """Model formulation"""
    prob = LpProblem("Deal Allocation Problem", LpMaximize)

    """Objective"""
    k=[]
    for i in range(num_of_purchasers):
        k.append( [deals[j][1] * x[(deals[j][0], i)] for j in range(deal_count)])
    prob += lpSum(k)

    """Hard Constraints"""
    '''Ensure that the Allocated amount is less that the Pending to Allocated amount'''
    for i in range(num_of_purchasers):
        if purchasers[i] > 0:
            prob += lpSum([deals[j][1] * x[(deals[j][0], i)] for j in range(deal_count)]) <= purchasers[i]*y[i]
        else:
            prob += lpSum([deals[j][1] * x[(deals[j][0], i)] for j in range(deal_count)]) == 0
            
    '''Ensure that a deal can only be allocated once'''
    for idx, item,in enumerate(deals):
        if deals[idx][1] > 0:
            prob += lpSum([x[(item[0], j)] for j in range(num_of_purchasers)]) <= 1
        else:
            prob += lpSum([x[(item[0], j)] for j in range(num_of_purchasers)]) == 0

    '''Ensure that each purchaser be allocated at least one deal'''
    if min_deal == True:    
        for i in range(num_of_purchasers):
            if purchasers[i] >= min_deal_value:
                m = []
                for j in range(deal_count):
                    if deals[j][1] > 0:
                        m.append(x[(deals[j][0], i)]) 
                prob += lpSum(m) >= 1

    """Soft Constraints"""   
    if pref_penalty:
        penalty = -10*100
        k_1 = []
        k_0 = []
        
        type_0, type_1 = [], []
        for idx, item,in enumerate(deals_type):
            if item[1] == 1:
                type_1.append((item, deals[idx]))
            elif item[1] == 0:
                type_0.append((item, deals[idx]))
                
        for i, item,in enumerate(allocation_pref):
            if item != 2 and purchasers[i] > 0:

                '''Add soft constraints for type 1 deals being 0 or greater than 1 (depending on pref)'''
                s = 1*allocation_pref[i] if purchasers[i] > 0 else 0
                k_1.append([1 * x[(type_1[j][0][0], i)] for j in range(len(type_1))])
                z_1 = [LpVariable(k_1[-1][j], lowBound = 0, upBound = 1, cat = pulp.LpInteger) for j in range(len(type_1))]
                # LpAffineExpression([(var1,coeff1), (var2,coeff2)])
                c_e_LHS = LpAffineExpression([(z_1[j], 1) for j in range(len(type_1))])                                         
                # The constraint LHS = RHS    #"sense=0" ==> "=="
                c_e_pre = LpConstraint(e=c_e_LHS, sense=s, name='elastic %s_%s' % (i,1), rhs=s)
                # The definition of the elasticized constraint
                c_e = c_e_pre.makeElasticSubProblem(penalty=penalty)
                # Adding the constraint to the problem
                prob.extend(c_e)

                '''Add soft constraints for type 0 deals being 0 or greater than 1 (depending on pref)'''
                s_0 = 0 if s == 1 else 1
                k_0.append([1 * x[(type_0[j][0][0], i)] for j in range(len(type_0))])
                z_0 = [LpVariable(k_0[-1][j], lowBound = 0, upBound = 1, cat = pulp.LpInteger) for j in range(len(type_0))]
                c_e_LHS = LpAffineExpression([(z_0[j], 1) for j in range(len(type_0))]) 
                c_e_pre = LpConstraint(e=c_e_LHS, sense=s_0, name='elastic %s_%s' % (i,0), rhs=s_0)
                c_e = c_e_pre.makeElasticSubProblem(penalty=penalty)
                prob.extend(c_e)
                

    try:
        ##  prob.solve(GLPK(options=['--mipgap', str(target_gap), "--tmlim", "1"]))                                         #Solve by minimum gap
        prob.solve(GLPK(options=["--tmlim", time_limit]))                                                                   #Solve by time limit
    except Exception as e:
        print(e)
        print("Status:", LpStatus[prob.status])
        print("Process stopped.")
        raise SystemExit(0)

    if prin:
        print(prob)
        print("Status:", LpStatus[prob.status])

    if LpStatus[prob.status] == "Infeasible":
        if not prin : print("Status:", LpStatus[prob.status])
        raise SystemExit(0)
    
    k = []
    for i in x.keys():
        if x[i].value() == 1:
            k.append(i)

    m, optimum_alloaction = format_output(k, deal_count, prin = prin)
    if prin: print(optimum_alloaction)
    return optimum_alloaction
        
if __name__ == "__main__":

    filename = "AllocationModelFinal.xlsm"
    ws_name = "Deals"
    time_limit = "60"
    prin = True
    
    '''Get inputs from Excel'''
    xl, wb, ws = get_excel(filename, ws_name)
    allocation_pref, deals, deals_type, purchasers, target_gap, min_deal, \
                     pref_penalty = get_vaiables(wb, ws, prin=prin)

    '''Test that lengths of key variables match.'''
    if len(deals) != len(deals_type):
        print("Not all deals have a defined type.")
        raise SystemExit(0)
    if len(allocation_pref) != len(purchasers):
        print("Not all purchasers have a defined preference.")
        raise SystemExit(0)
    min_deal_value = 10**9  #random large number
    for idx, item in enumerate(deals):
        if min_deal_value > item[1] and item[1] != 0:
            min_deal_value = item[1]
    if len(deals) != len(ws.Range("Optimum_Alloaction").Value):
        print("Output length does not match number of deals.")
        raise SystemExit(0)
    
     
    '''Optimize deal allocation'''
    optimum_alloaction = optimize_deal_allocation(deals,
                                                  purchasers,
                                                  deals_type,
                                                  allocation_pref,
                                                  target_gap,
                                                  min_deal,
                                                  pref_penalty,
                                                  time_limit,
                                                  min_deal_value,
                                                  prin=prin)

    '''Put the optimum_alloaction in the Excel.'''
    put_optimum_alloaction(ws, optimum_alloaction)

