case (TaxState contains "MD"  OR TaxState contains "MA") AND Order≠"" 
    vMDMAedible=""
    ;arrayfilter(vMDMAedible,vMDMAedible,¶,?(val(array(vMDMAedible,¶,seq(),))))
            //****selects the Amounts for all orders over 4700
    ;arrayfilter Order, taxitems,¬,?(val(extract(extract(Order,¶,seq()),¬,2))>4700,extract(extract(Order,¶,seq()),¬,7),"")
            //TEST: Get range 4700-5931, exclude 5932-5939,get above 5940
   
   //gives a list of item numbers in Order
   vArraySize=arraysize(Order,¶) //counts number of different items in Order
   vPosIncrement=1
   vCheckArray=""
   vMDMAedible=""
   loop
    vCheckArray=?(val(striptonum(array(array(Order,vPosIncrement,¶),2,¬)))>4700,striptonum(array(array(Order,vPosIncrement,¶),2,¬))+¬+array(array(Order,vPosIncrement,¶),7,¬),"") //grabs the item nums and their price if the item  number is more than 4700
        if val(vCheckArray[1,¬])≤5932 and val(vCheckArray[1,¬])>4700
        vMDMAedible=array(array(Order,vPosIncrement,¶),7,¬)+vMDMAedible
        else 
        if val(vCheckArray[1,¬])>4700
        vMDMAedible=array(array(Order,vPosIncrement,¶),7,¬)+vMDMAedible
        endif
        endif
        vPosIncrement=vPosIncrement+1
   until vPosIncrement=vArraySize+1
    ;arraysort vMDMAedible,vMDMAedible,"^" //puts them in acensding order
   message vMDMAedible
   clipboard()=vMDMAedible
    stop
    ;taxitems=arraystrip(taxitems,¬)
    ;TaxTotal=arraynumerictotal(taxitems,¬)
    ;TaxedAmount=?(Taxable="Y", TaxTotal*float(divzero(AdjTotal,Subtotal)),0)