local taxstates, noship, taxitems, vgroupvermonttax, vMDMAedible, vArraySize, vPosIncrement, vCheckArray
taxstates="CT,GA,IL,IN,KS,KY,MI,MN,NC,NJ,NM,NY,OH,PA,RI,TN,VT,WA,WI,WV,WY"
noship="FL,MA,MD,ME,VA,UT" ; IA removed 11-18-19

;for orders with taxable shipping and not taxing edibles
case (TaxState contains "VT" OR TaxState contains "RI" OR TaxState contains "CT") AND Order≠""   
    arrayfilter Order, taxitems,¬,?(val(extract(extract(Order,¶,seq()),¬,2))>4700,extract(extract(Order,¶,seq()),¬,7),"")
    taxitems=arraystrip(taxitems,¬)
    TaxTotal=arraynumerictotal(taxitems,¬)
    TaxedAmount=?(Taxable="Y", TaxTotal*float(divzero(AdjTotal,Subtotal))+float(«$Shipping»*float(divzero(TaxTotal,Subtotal))),0)
local taxstates, noship, taxitems, vgroupvermonttax, vMDMAedible, vArraySize, vPosIncrement
taxstates="CT,GA,IL,IN,KS,KY,MI,MN,NC,NJ,NM,NY,OH,PA,RI,TN,VT,WA,WI,WV,WY"
noship="FL,MA,MD,ME,VA,UT" ; IA removed 11-18-19

;for orders with taxable shipping and not taxing edibles
case (TaxState contains "VT" OR TaxState contains "RI" OR TaxState contains "CT") AND Order≠""   
    arrayfilter Order, taxitems,¬,?(val(extract(extract(Order,¶,seq()),¬,2))>4700,extract(extract(Order,¶,seq()),¬,7),"")
    taxitems=arraystrip(taxitems,¬)
    TaxTotal=arraynumerictotal(taxitems,¬)
    TaxedAmount=?(Taxable="Y", TaxTotal*float(divzero(AdjTotal,Subtotal))+float(«$Shipping»*float(divzero(TaxTotal,Subtotal))),0)

;for group cover sheets with taxable shipping and not taxing edibles
case (TaxState contains "VT" OR TaxState contains "RI" OR TaxState contains "CT") AND Order="" AND Subtotal>0
    arrayselectedbuild vgroupvermonttax,",","", ?(str(OrderNo) contains ".",?(TaxTotal>0,TaxTotal,TaxedAmount),"")
    vgroupvermonttax=arraynumerictotal(vgroupvermonttax,",")
    TaxTotal= vgroupvermonttax
    TaxedAmount=?(Taxable="Y", float(TaxTotal*float(1-Discount-?(MemDisc>0,.01,0))+float(«$Shipping»)*float(TaxTotal)/Subtotal),0)

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

;for group cover sheets not taxing shipping and not taxing edibles
case (TaxState contains "MD" OR TaxState contains "MA") AND Order="" AND Subtotal>0
    arrayselectedbuild vgroupvermonttax,",","", ?(str(OrderNo) contains ".",?(TaxTotal>0,TaxTotal,TaxedAmount),"")
    vgroupvermonttax=arraynumerictotal(vgroupvermonttax,",")
    TaxTotal= vgroupvermonttax
    TaxedAmount=?(Taxable="Y", float(TaxTotal*float(1-Discount-?(MemDisc>0,.01,0))),0)

case arraycontains(taxstates+","+noship,TaxState,",")=0 
    TaxedAmount=0

case arraycontains(noship,TaxState,",")=-1 
    TaxedAmount=?(Taxable="Y",AdjTotal,0)

case arraycontains(taxstates,TaxState,",")=-1 
    TaxedAmount=?(Taxable="Y",AdjTotal+Surcharge+«$Shipping»,0) 

endcase 

if str(OrderNo) notcontains "."
    SalesTax=round(float(TaxedAmount)*float(TaxRate)+.001,.01) 
    StateTax=round(float(TaxedAmount)*float(StateRate)+.001,.01) 
    CountyTax=round(float(TaxedAmount)*float(CountyRate)+.001,.01) 
    CityTax=round(float(TaxedAmount)*float(CityRate)+.001,.01) 
    SpecialTax=round(float(TaxedAmount)*float(SpecialRate)+.001,.01)
endif


    TaxedAmount=0

case arraycontains(noship,TaxState,",")=-1 
    TaxedAmount=?(Taxable="Y",AdjTotal,0)

case arraycontains(taxstates,TaxState,",")=-1 
    TaxedAmount=?(Taxable="Y",AdjTotal+Surcharge+«$Shipping»,0) 

endcase 

if str(OrderNo) notcontains "."
    SalesTax=round(float(TaxedAmount)*float(TaxRate)+.001,.01) 
    StateTax=round(float(TaxedAmount)*float(StateRate)+.001,.01) 
    CountyTax=round(float(TaxedAmount)*float(CountyRate)+.001,.01) 
    CityTax=round(float(TaxedAmount)*float(CityRate)+.001,.01) 
    SpecialTax=round(float(TaxedAmount)*float(SpecialRate)+.001,.01)
endif

