import pandas 
df_vpr=pandas.read_excel(open('C:\Users\Charizard\Desktop\New folder\custody_vpr.xlsx','rb'), sheetname='Sheet1')
df_vpr_sort=df_vpr.sort(['Custody'])
counter=0
for index, row in df_vpr_sort.iterrows():
    if counter==0:
        df_vpr_sort.set_value(index, 'Group', 1, takeable=False)
        prev_Custody=row.Custody
        prev_Paydate=row.Paydate
        counter=counter+1
        Group_number=1
    if counter==1:
        if prev_Custody==row.Custody and prev_Paydate==row.Paydate:
            df_vpr_sort.set_value(index, 'Group', Group_number, takeable=False)
        else:
            Group_number=Group_number+1
            df_vpr_sort.set_value(index, 'Group', Group_number, takeable=False)
            prev_Custody=row.Custody
            prev_Paydate=row.Paydate
max_group=Group_number
df_mapping=pandas.read_excel(open('C:\Users\Charizard\Desktop\New folder\mapping.xlsx','rb'), sheetname='Sheet1')
df_mapping=df_mapping.fillna('')

for index, row in df_vpr_sort.iterrows():
    select_Custody=row.Custody
    select_Paydate=row.Paydate
    select_Group_number=str(row.Group)
    for index, row in df_mapping.iterrows():
        
        if row.Custody==select_Custody:
            df_mapping.set_value(index, 'Group', row.Group+','+select_Group_number, takeable=False)
            prev_group_number=select_Group_number
#delete mapping file rows where group number not present
            
df_ole=pandas.read_excel(open('C:\Users\Charizard\Desktop\New folder\ole_output.xlsx','rb'), sheetname='Sheet1')
df_ole['Group']='0'
for index, row in df_mapping.iterrows():
    wins=row.Wins_Main
    select_Group_number=row.Group
    
    for index,row in df_ole.iterrows():
        if row.Wins_Main==wins:
            df_ole.set_value(index, 'Group', row.Group+select_Group_number, takeable=False)
amount_wins=0
amount_vpr=0
counter_port=0
df_vpr_sort['Amount_matching'] ='Add comment' 
for i in range(1,max_group+1):#one to one case case
    j=str(i)
    
            
     
    for index, row in df_vpr_sort.iterrows():
        row.Group
        group_string=str(row.Group)
        group_list=group_string.split(',')
        
        if j in group_list:
            print "done"
            index_vpr=index
            counter_port=counter_port+1
            amount_vpr=row.Reclaims_amount
            for index,row in df_ole.iterrows():
                group_string=str(row.Group)
                group_list=group_string.split(',')
                if  j in group_list :
                    amount_wins=row.Amount
                    if amount_vpr==amount_wins:
                        df_vpr_sort.set_value(index_vpr,'Amount_matching' ,'One to one match', takeable=False)
                        
                        break
                    
                    else:
                        continue
amount_vpr=0
amount_wins=0
 #multi vpr
for i in range(1,max_group+1):
    j=str(i)
    print i
    amount_vpr=0
    amount_wins=0
    for index, row in df_vpr_sort.iterrows():
        group_string=str(row.Group)
        print group_string
        
        group_list=group_string.split(',')
        if j in group_list and row.Amount_matching<>'One to one match' :
            index_vpr=index
            counter_port=counter_port+1
            amount_vpr=amount_vpr+row.Reclaims_amount
            df_vpr_sort.set_value(index, 'Amount_matching', "In Process", takeable=False)
            print amount_vpr
    for index,row in df_ole.iterrows():
        group_string=str(row.Group)
        group_list=group_string.split(',')
        if j in group_list:
            index_vpr=index
            counter_port=counter_port+1
            amount_wins=amount_wins+row.Amount
            print amount_wins
    if amount_wins==amount_vpr:
        for index, row in df_vpr_sort.iterrows():
            if row.Amount_matching=='In Process':
                df_vpr_sort.set_value(index, 'Amount_matching', "Multi port match", takeable=False)
    else:
        for index, row in df_vpr_sort.iterrows():
            if row.Amount_matching=='In Process':
                df_vpr_sort.set_value(index, 'Amount_matching', "Add comment", takeable=False)
        
        
        
        
    
        
    
        
    
        
        
    
            
            
    
            
    
    

            
