˛.InitializeˇExpressionStackSize 250000000
global n, raya, rayb, rayc, rayd, raye, rayf, rayg, rayh, rayi, rayj, waswindow, oldwindow,
Com, Cost, Disc, VDisc, ODisc, s
, item, size, ono, vfax, state, sub, vzip, zprintedgroups,
vd, adj, tax, f, gr, di, da, three, secno, alt, Numb, vchecker, Vship, form, zitem, zsize,
zcomment, linenu, zline, order, zprice, zfill, ztot, addon, brange, back, status, backline,
added, vcode, oldfile, pickno, members, origin, each, vmail, vupdate, zname, subnum, atotal,
zselect, vearly, zemail, folder, file, type, zpuller, zchecker, data, zlarge, zcancel, zprintrun,
zorderarray, zordercount, zcode, zmanifest, zshipment, zdun, vmem, vmemflag, zstat, zdiagnose,
zflatrate, vmailbatch, zbackrun, zlinect, zleftitem, zleftbo, zleftct, zcheckmin, zyear, zcomyear,
zcanada, zseedwt, zpacketwt, zshipwt, zcanadaorders, zbackrunning, ztime, ztimedif, zlevel,
zmemEmail, zmemPhone, zmemCon, zmemAdd, zmemNo, zmemGroup, zmembers, zmemberwindow, zadjlist,
zdiscEmail, zdiscPhone, zdiscCon, zdiscAdd, zdiscNo, zdiscGroup, zdiscounts, zdiscountwindow, zstatuses,
zmethod, zpayfield, vfilled, ztaxcase, ztaxstates, zCCrefunds, zrecall
zcanada="No"
zyear="44" ;this changes the year prefix for all but comments at the beginning of a new fiscal year and ship lookup files (for ogs items), 
//also does not work for ImportOrders applescript
zcomyear="44" ;this changes the year prefix for comments and ship lookup files which don't change in July



zprintedgroups=""

;********This is for Tax Reporting*******    
case folderpath(dbinfo("folder","")) CONTAINS "wayfair"
    save
    goto Skip
endcase

if folderpath(dbinfo("folder","")) CONTAINS "Facilitator" ;need tweaking for facilitator copy as they drop the leading space - see below as well
openform "seedspagecheck"
    zoomwindow 23,9,710,620,""
arraybuild zprintedgroups,¶,"", ?(group_count>0,str(OrderNo)[1,6],"") ;changed from int(OrderNo) to see if group_count on .001 picksheet would print actual number of pieces in Group
else

if folderpath(dbinfo("folder","")) CONTAINS "register"
waswindow= info("WindowName")
message "This updates the customer_list file automatically. It can take a while. Wait for the next message."
NoYes "Do you still want to update the customer_list file?"
if clipboard()="Yes"
OpenFile "customer_list"
openfile "&&"+zcomyear+"seedstally"
goform "order_finder"
zoomwindow 500,5,200,1145,""
save
window waswindow
message "The customer_list is now up-to-date."
endif
save
closefile
stop
endif
if folderpath(dbinfo("folder","")) CONTAINS "MAIL"
NoYes "Do you need the customer record open?"
goform "shipping_form"
call export_mail/µ
save
openanything "","Endicia alias"
if clipboard()="No"
quit
else
vmail=""
stop
endif
endif
Numb=0
members=""
addon=""
atotal=0
brange=""
n=1
raya=¶
rayb=""
rayc=""
rayd=""
raye=""
rayf=""
rayg=""
rayh=""
rayi=""
rayj=""
vchecker=""
vmail=""
zemail=""
zcancel="No"
ODisc=0
vmem=0
vmemflag=0 ;this was set to handle group members & non-members - not in use
zstat=""
zdiagnose=""
zbackrun="Yes"
zcheckmin=2
zseedwt=0
zpacketwt=0
zshipwt=0
zcanadaorders=""
zbackrunning="No"
ztime=now()
ztimedif=180
zlevel=0
zstatuses=""
zmethod="Cash"
ztaxstates="CT,GA,IL,IN,KS,KY,MI,MN,NC,NJ,NM,NY,OH,PA,RI,TN,VT,WA,WI,WV,WY,FL,MA,MD,ME,VA,UT" ;IA-removed 1/20
zCCrefunds="N"
zrecall="N"
selectall
if folderpath(dbinfo("folder","")) CONTAINS "reporting"
stop
endif
if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
waswindow=zyear+"seedstally:seedspagecheck"
if info("Windows") notcontains zyear+"seedstally:seedspagecheck"
WindowBox "23 9 733 630"
openform "seedspagecheck"
else
window zyear+"seedstally:seedspagecheck"
zoomwindow 23,9,710,620,""
endif
if info("Windows") notcontains "ViewBackorder"
WindowBox "30 640 600 1120 noPalette noHorzScroll"
openform "ViewBackorder"
else
goform "ViewBackorder"
zoomwindow 30,640,550,490,"noPalette noHorzScroll"
endif
if info("Windows") notcontains "ViewBackHistory"
WindowBox "580 640 860 1130 noPalette noHorzScroll"
openform "ViewBackHistory"
else
goform "ViewBackHistory"
zoomwindow 580,640,280,490,"noPalette noHorzScroll"
endif
openfile zyear+"RefundRegister"
ZoomWindow 746,9,110,620,""
openfile "MailManifest"
if info("Records") >1
Message "There are already records in the Mail Manifest. Make sure these don't have to be cleared out before you start running paperwork. Thanks."
endif
OpenFile zcomyear+"SeedsComments"
window "Hide This Window"
;OpenFile zcomyear+"canada_weights"
;window "Hide This Window"
openfile "pullers&checkers"
window "Hide This Window"
if zCCrefunds="Y"
OpenFile "credit charges"
ZoomWindow 746,9,110,620,""
endif
window waswindow
Field OrderNo
sortup
firstrecord
save
;message "Start order after inventory is 13654 & 90545. Date 5-11-17"
;message "still need to clear & link tally, link register and test"
if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
message "ready to run paperwork"
endif
endif
endif ;part of facilitator pagecheck view


Skip:
save

gosheet
˛/.Initializeˇ˛.commentsˇglobal vOrderNo
zcanada="No"
zcanadaorders=""
    Case form="seedspagecheck"
    openfile "SeedsTotaller"
    endcase
oldwindow=info("windowname")
window waswindow
firstrecord
loop
vOrderNo=OrderNo
vmem=0
        if Order contains "201)"
        zlarge=zlarge+1
        endif
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            zcanadaorders=?(zcanadaorders="",str(OrderNo),zcanadaorders+", "+str(OrderNo))
            field Zip
            clearcell
            endif
    if PickSheet=""
    order=Order
        if Subs=""
        Subs=?(Sub1="Y" and Sub2="Y","Y",?(Sub1="Y" and Sub2="N","O",?(Sub1="N" and Sub2="Y","C","N")))
        endif
    sub=Subs
    ono=OrderNo
if MemDisc>0 OR Discount=.5
vmem=1
endif
    window oldwindow
    openfile "&@order"
    call ".commentfill"

        if str(OrderNo) contains ".001"
    call ".grp_count"
        endif
        
    endif
;stop    
downrecord
;stop
zcanada="No"
until info("stopped")
˛/.commentsˇ˛.grp_countˇarrayselectedbuild zprintrun, ",", "",str(OrderNo)[1,6] ; build a list of all selected order numbers
arrayfilter zprintrun, zorderarray, ",", ?(import()=str(OrderNo)[1,6],str(OrderNo)[1,6],"") ;strip to the order being worked on
zorderarray=arraystrip(zorderarray, ",") ;strip out blank elements
zordercount=arraysize(zorderarray,",")
group_count=zordercount
;message zordercount

;arrayselectedbuild zprintrun, ",", "",int(OrderNo) ;6-2 change ; build a list of all selected order numbers
;arrayfilter zprintrun, zorderarray, ",", ?(import()=int(OrderNo),int(OrderNo),"") ;strip to the order being worked on ;6-2 change
;zorderarray=arraystrip(zorderarray, ",") ;strip out blank elements
;zordercount=arraysize(zorderarray,",")
;group_count=zordercount
;message zordercount˛/.grp_countˇ˛.retotalˇGrTotal=0
;if  (OrderNo>10000 AND OrderNo<30000) OR (OrderNo>70000 AND OrderNo<100000)
;if Pool≠2       ;pool 2 designation is for orders with a set amount for shipping
;«$Shipping»=0
;endif
;endif

    case OrderNo<30000 OR (OrderNo>70000 AND OrderNo<1000000) ; 6-2 change

;for group pieces
                if OrderNo≠int(OrderNo) ;6-2 change
                    MemDisc=?(MemDisc>0, float(Subtotal-OGSTotal)*float(.01),0)
                    AdjTotal=Subtotal-MemDisc
                    SalesTax=0
;                    SalesTax=?(Taxable="Y", float(AdjTotal)*float(.055),0) ;sales tax is now determined by the group coordinators's tax status and delivery address
                    call ".salestax"
                    OrderTotal=0
                    Patronage=0
                    RealTax=0
                    Paid=0
                    «1stPayment»=0
                    «1stTotal»=0
                    «1stRefund»=0
                    «BalDue/Refund»=0
                    VolDisc=0
                    Discount=0
               Goto end
                endif
                
;now do regular orders and group cover sheets
        VolDisc=(float(Subtotal-OGSTotal)*float(Discount))
        
;David and I agreed to not cover the unlikely case that the entire group does not get the member discount 12-26-12
;meaning all groups with any member discount will get 1% at this point - sales tax will also be adjusted similarly
;                if vmemflag=0      ;does not calculate 1% member discount when group pieces are members individually and the group is not a member.
                MemDisc=?(MemDisc>0, float(Subtotal-OGSTotal)*float(.01),0)
;                endif             
    ;but avoid adjusting tax total on groups - it's adjusted further down                    
    ;If Order notcontains "0000" AND Order ≠ "" AND (Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderComments CONTAINS "empty order")
    ;AND (Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderNo>99000)
    ;TaxTotal=?(MemDisc>0, float(TaxTotal)*float(.99),TaxTotal)
    ;endif     
                organic_total=?(MemDisc>0,float(organic_subtotal)*(1-(Discount+.01)),float(organic_subtotal)*(1-Discount)) 
       
        AdjTotal=Subtotal-VolDisc-MemDisc

if zcanada="No"
call ".zone"
endif

zshipwt=float((«1SpareNumber»+«2SpareNumber»)/28.35) ;convert seed+packet weight to ounces
;message zshipwt
;now add in packaging
    if zcanada="Yes"
    zshipwt=?(zshipwt≤15 AND zshipwt>0, zshipwt+1,?(zshipwt≤64 AND zshipwt>0, zshipwt+5,?(zshipwt>240,zshipwt+16,zshipwt)))
    else
    zshipwt=?(zshipwt≤15 AND zshipwt>0, zshipwt+1,?(zshipwt>240,zshipwt+16,zshipwt))
    endif
3SpareNumber=round(zshipwt,.01)

            if ShippingWt>0 AND Pool≠2 AND zcanada="No" ;pool 2 allows a fixed amount for shipping
            call ".shipping"
            endif
            
            if zcanada="Yes" AND Pool≠2 ;pool 2 allows a fixed amount for shipping
            «$Shipping»=0
                if ShipCode ≠ "P"
            call ".shipping_canada"
                endif
            endif

     if Subtotal=0
     «$Shipping»=0
     endif   

call ".salestax"

        OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»+AddOns
        GrTotal=OrderTotal+Donation+Membership
        
     endcase
     vmemflag=0
    end:˛/.retotalˇ˛.salestaxˇlocal taxstates, noship, taxitems, vgroupvermonttax
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

;for orders not taxing shipping and not taxing edibles
case (TaxState contains "MD"  OR TaxState contains "MA") AND Order≠"" 
    arrayfilter Order, taxitems,¬,?(val(extract(extract(Order,¶,seq()),¬,2))>4700,extract(extract(Order,¶,seq()),¬,7),"")
    taxitems=arraystrip(taxitems,¬)
    TaxTotal=arraynumerictotal(taxitems,¬)
    TaxedAmount=?(Taxable="Y", TaxTotal*float(divzero(AdjTotal,Subtotal)),0)

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

˛/.salestaxˇ˛.order_checkerˇsynchronize
selectall
;stop

;case ztaxcase="me"
;select TaxRate≠StateRate+CountyRate+CityRate+SpecialRate and TaxState contains "ME"
;     and Taxable notcontains "N"


case ztaxcase="pre"
select TaxState="" AND ShipCode notcontains "D" AND «9SpareText»=""
selectadditional  (ztaxstates contains TaxState AND TaxRate=0) AND Taxable contains "Y"
    AND str(OrderNo) notcontains "." AND ShipCode notcontains "D"
selectadditional  (ztaxstates notcontains TaxState AND TaxRate≠0) AND Taxable contains "Y"
    AND str(OrderNo) notcontains "." AND ShipCode notcontains "D"
selectadditional TaxState="IL" AND TaxRate≠.0625 AND (TaxRate≠0 AND str(OrderNo) contains ".")
selectadditional TaxState="PA" AND TaxRate≠.06 AND (TaxRate≠0 AND str(OrderNo) contains ".")
selectadditional Taxable=""
selectadditional TaxRate≠0 AND Taxable contains "Y" AND str(OrderNo) contains "."
selectadditional Order CONTAINS "1)"+¬+"0"+¬+¬+"0"
 
;selectwithin OrderPlaced<date(|||10/01/21|||)
;case ztaxcase="post"
;select SalesTax≠StateTax+CountyTax+CityTax+SpecialTax and Status≠"" 
;    and abs((StateTax+CountyTax+CityTax+SpecialTax)-SalesTax)>.01
    
endcase

ztaxcase=""˛/.order_checkerˇ˛.refundˇzmethod="Cash"
data= info("DatabaseName") 
If «BalDue/Refund» ≥ zcheckmin OR (zcanada="Yes" AND «BalDue/Refund» > .10) OR vmail="pBm"
    if zCCrefunds="N" or CreditCard="" or OrderNo<710000
    YesNo "Do you want to write a check for "+pattern(«BalDue/Refund», "$#.##")+"?"
        if clipboard()="Yes"
        zmethod="Check"
       endif
       else
       YesNo "Do you want to refund the credit card for "+pattern(«BalDue/Refund», "$#.##")+"?"
        if clipboard()="Yes"
        zmethod="Credit Card"
       endif
       endif

        if clipboard()="No"
        message "Refund not complete. You haven't chosen if you want to give refund as cash, check, or credit card. Please rerun this when you've decided."
        stop
        endif
call .RefundRecord

if zmethod="Credit Card"
global raycc
waswindow=info("windowname")
arraylinebuild raycc,¶,"",¬+str(OrderNo)+¬+str(Zip)+¬+Email+¬+str(GrTotal)+¬+str(abs(«BalDue/Refund»))+¬+str(«Auth_Code»)+¬+CreditCard+¬+ExDate
window "credit charges"
openfile "+@raycc"
window waswindow
endif


if zmethod="Check"
        waswindow=info("windowname")
        window zyear+"RefundRegister"
        insertbelow
        Name=grabdata(data, Group)+" or "+grabdata(data, Con)
        Name=?(grabdata(data, Group)="", grabdata(data, Con),Name)
        «$Amount»=grabdata(data, «BalDue/Refund»)
         Amount=grabdata(data, «BalDue/Refund»)
       Date=today()
       OrderNO=grabdata(data, OrderNo)
       Con=grabdata(data, Con)
       Group=grabdata(data, Group)
       MAd=grabdata(data, MAd)
       City=grabdata(data, City)
       St=grabdata(data, St)
       Zip=grabdata(data, Zip)
        OpenForm "Check2"
        PrintOneRecord
        CloseWindow
        FirstRecord
        LastRecord
        Save
        window waswindow

        ;the Morse Pitts warning
        if «BalDue/Refund» ≥ 2000
        message "Please warn Gene about the size of this check so we're ready at the bank! Thanks."
        endif

       ;else 
       ;message "Don't forget to give a CASH REFUND of "+pattern(«BalDue/Refund», "$#.##")+"."

endif
else 
call .RefundRecord
message "Don't forget to give a CASH REFUND of "+pattern(«BalDue/Refund», "$#.##")+"."         
endif
˛/.refundˇ˛.RefundRecordˇcase AddPay1=0
    AddPay1=0-«BalDue/Refund»
    MethodPay1=zmethod
    DatePay1=today()
case AddPay2=0
    AddPay2=0-«BalDue/Refund»
    MethodPay2=zmethod
    DatePay2=today()
case AddPay3=0
    AddPay3=0-«BalDue/Refund»
    MethodPay3=zmethod
    DatePay3=today()
case AddPay4=0
    AddPay4=0-«BalDue/Refund»
    MethodPay4=zmethod
    DatePay4=today()
case AddPay5=0
    AddPay5=0-«BalDue/Refund»
    MethodPay5=zmethod
    DatePay5=today()
case AddPay6=0
    AddPay6=0-«BalDue/Refund»
    MethodPay6=zmethod
    DatePay6=today()
case AddPay6≠0
 message "This order has more than 6 payment adjustments and needs to be adjusted manually."
stop
endcase˛/.RefundRecordˇ˛.baldueˇif zcanada="No"
message "Don't forget the BALANCE DUE of "+pattern(abs(«BalDue/Refund»), "$#.##")
endif
if zcanada="Yes"
message "Don't forget the BALANCE DUE of "+pattern(abs(«BalDue/Refund»), "$#.##")+". For Canadian orders, please check with Ryan or Bernice before this is shipped, thanks."
endif
Donated_Refund=0
˛/.baldueˇ˛.addpayˇlocal zpayment
zpayment=0
zpayfield=info("fieldname")
case zpayfield="AddPay1"
    zpayment=AddPay1
    DatePay1=?(AddPay1≠0,today(),0)
    if AddPay1≠0
        field MethodPay1
        editcell
    else
        MethodPay1=""
    endif

case zpayfield="AddPay2"
    zpayment=AddPay2
    DatePay2=?(AddPay2≠0,today(),0)
    if AddPay2≠0
        field MethodPay2
        editcell
    else
        MethodPay2=""
    endif

case zpayfield="AddPay3"
    zpayment=AddPay3
    DatePay3=?(AddPay3≠0,today(),0)
    if AddPay3≠0
        field MethodPay3
        editcell
    else
        MethodPay3=""
    endif

case zpayfield="AddPay4"
    zpayment=AddPay4
    DatePay4=?(AddPay4≠0,today(),0)
    if AddPay4≠0
        field MethodPay4
        editcell
    else
        MethodPay4=""
    endif

case zpayfield="AddPay5"
    zpayment=AddPay5
    DatePay5=?(AddPay5≠0,today(),0)
    if AddPay5≠0
        field MethodPay5
        editcell
    else
        MethodPay5=""
    endif

case zpayfield="AddPay6"
    zpayment=AddPay6
    DatePay6=?(AddPay6≠0,today(),0)
    if AddPay6≠0
        field MethodPay6
        editcell
    else
        MethodPay6=""
    endif

endcase
Paid=Paid+zpayment
«BalDue/Refund»=Paid-GrTotal
if Paid≠«1stPayment»+AddPay1+AddPay2+AddPay3+AddPay4+AddPay5+AddPay6
message "The individual payment amounts do not add up to what's in Paid. Please be sure this is correct. Talk to me if you're unsure, it can avoid bookkeeping headaches. Thanks. Ryan"
endif
˛/.addpayˇ˛.doneˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
zcanada="No"
status=Status
form=info("FormName")
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
    if PickSheet="" AND Notes1 NOTCONTAINS "Order cancelled"
    if  info("Selected") = info("Records") 
    message "This picksheet's never been run or you may need to synchronize. Check it out."
    stop
    endif
    endif
waswindow=info("windowname")

if Order contains "5965" and Notes1 notcontains "soil kit included"
Message "Please include soil test kit, indicate soil kit included with the order and then finish running the paperwork."
stop
endif

;this sets group coversheet status depending on the status of the groups pieces
zstatuses=""
if vmail="BG" and Order=""
arrayselectedbuild zstatuses,",","",?(str(OrderNo) contains ".",Status,"")
;message zstatuses
if zstatuses notcontains "B"
settrigger "Button.Complete"
else
settrigger "Button.BackOrder"
endif
;message info("trigger")
endif

            if vmail="" and str(OrderNo) notcontains "." and zcancel≠"Yes"
            YesNo "Does this order need a mail flag?"
            if clipboard()="Yes"
;                loop
                GetScrapOK "A +/or B +/or C?"
;                                    if clipboard() = ""
;                                    message "Please enter a mail flag"
;                                    endif
;                until clipboard() ≠ ""                
                vmail=upper(clipboard())
                    if clipboard() CONTAINS "x" OR clipboard() CONTAINS "p"
                    MailFlag=clipboard()[1,1]+upper(clipboard()[2,-1])
                    else
                    MailFlag=upper(clipboard())
                   endif
                   if MailFlag contains "B"
                        noyes "Does this order need to be added to the mail manifest"
                        if clipboard()="Yes"
                        vmail=MailFlag
                        call ManifestThisOrder
                        endif
                   endif
            endif
            endif
if vmail="BG" and str(OrderNo) notcontains "." 
MailFlag="B"
endif
if zstat="" AND vmail NOTCONTAINS "B"
    if Checker="Nobody" OR Checker="" ;6-2 change        
         if str(OrderNo) NOTCONTAINS "."
             NoYes  "There's no checker. Do you want to fix it?"
                 If clipboard()="Yes"
call .checkers
    stop
    endif
        endif
    endif
    if (Puller = "Everybody" OR Puller = "") 
        if  info("Selected") = info("Records") 
    NoYes  "There's no puller. Do you want to fix it?"
    If clipboard()="Yes"
call .pullers
    stop
    endif
        endif
    endif
endif

if Backorder≠"" and (vmail contains "B" or Backorder NOTCONTAINS ¬+"  0"+¬ )
call .backorder_count
if zbackrun="Yes"
call .backorder_record
endif
endif

if info("trigger")="Button.Complete" OR (Order notcontains "backorder" AND Order≠"" AND Order notcontains "1)"+¬+"0"+¬+¬+"0") OR zcancel="Yes"

;do group pieces
    if str(OrderNo) CONTAINS  "."
        if Order CONTAINS "backorder"
        Message "This order still contains backorders! Check it out."
        stop
        endif
    Status="Com"
        if FirstFillDate=0
        FirstFillDate=today()
        endif
    if zrecall="N"    
    FillDate=today()
    endif
    Backorder=""
    if vmail notcontains "g"
    downrecord
    else
    call ".session"
    endif
        if str(OrderNo) contains "."
        if Status=""
    call .pullers
        endif
    else
    UpRecord
        endif
    zcancel="No"
    stop
    endif

;now everything else
if Status="Com"
    if vmail="pBm"
call .manifest
    endif
    if vmail="pB"
    call .add_to_endicia
    endif
goto Additional
endif

    if Order CONTAINS "backorder"
    Message "This order still contains backorders! Check it out."
    stop
    endif

    if vmail="pBm"
call .manifest
    endif
    if vmail="pB"
    call .add_to_endicia
    endif

Status="Com"
        if FirstFillDate=0
        FirstFillDate=today()
        endif
Additional:
if zrecall="N"
FillDate=today()
endif
Backorder=""
call ".retotal"

;resetting the donated refund
«BalDue/Refund»=Paid-GrTotal
    if «BalDue/Refund»=0
    Donated_Refund=0
    endif
if «BalDue/Refund»>0
    if «Donated_Refund»>0
       If «Donated_Refund»≥ «BalDue/Refund»
       «Donated_Refund»= «BalDue/Refund» 
       Paid=Paid- «BalDue/Refund» 
       «BalDue/Refund»=0
        endif
        If «Donated_Refund»<«BalDue/Refund»
        Paid=Paid- «Donated_Refund»
         «BalDue/Refund» =«BalDue/Refund»-«Donated_Refund»
         endif
         endif
            If «Donated_Refund»>(.2*OrderTotal) AND «Donated_Refund»>7
            YesNo "Nibezun is getting more than 20% of this order. OK?"
                if clipboard()="No"
                stop
                endif
            endif
 if «BalDue/Refund»=0
 goto skip   
 endif    
call ".refund"

skip:
Paid=Paid-«BalDue/Refund»
«BalDue/Refund»=0
endif
    If «BalDue/Refund»≥-1.50 AND «BalDue/Refund»<0
    Donated_Refund=0
    endif
If «BalDue/Refund»<-1.50
call ".baldue"
endif
    if Donated_Refund>0
    message "They've donated "+pattern(abs(Donated_Refund), "$#.##"+ " to Nibezun.")
    endif
RealTax=SalesTax+AddOnTax
Patronage=OrderTotal-RealTax
endif

if info("trigger")="Button.BackOrder" OR Order contains "backorder"

;do the group pieces
        if str(OrderNo) CONTAINS  "."
            if Order CONTAINS "backorder"
            Status="B/O"
        if FirstFillDate=0
        FirstFillDate=today()
        endif
            FillDate=today()
            call .backorder
    if vmail notcontains "g"
    downrecord
    else
    call ".session"
    endif
                if str(OrderNo) contains "."
        if Status=""
    call .pullers
        endif
            else
            uprecord
                endif
            zcancel="No"
           stop
            else
            Message "This order has no backorders! Check it out."
            zcancel="No"
            stop
            endif
            endif

;now group cover sheets
if  info("Selected") < info("Records")
    if vmail="pBm"
call .manifest
    endif
    if vmail="pB"
    call .add_to_endicia
    endif
                Status="B/O"
        if FirstFillDate=0
        FirstFillDate=today()
        endif
                FillDate=today()
                goto jump
                endif

;then everybody else
    if Order CONTAINS "backorder"
    if vmail="pBm"
call .manifest
    endif
    if vmail="pB"
    call .add_to_endicia
    endif
    Status="B/O"
        if FirstFillDate=0
        FirstFillDate=today()
        endif
    FillDate=today()
    call .backorder
    else
        Message "This order has no backorders! Check it out."
    stop
    endif

            jump:
call ".retotal"
«BalDue/Refund»=Paid-GrTotal
    If Paid=0
    call ".baldue"
    endif
endif

    if  info("Selected") < info("Records") AND zcancel≠"Yes"
    YesNo  "Do you want to print a cover sheet?"
    If clipboard()="Yes"
        if form="seedspagecheck"
;        GoForm "seedscoversheet"
        PrintUsingForm "", "seedscoversheet"
        PrintOneRecord
;        goform waswindow[":",-1][2,-1]
        endif
    endif
    
if ShipCode contains "P" AND OrderPickedUp=""
message "This pickup is supposed to be on the dock."
endif
    endif
    
rayg=""
members=""

if ShipCode contains "H" ;AND vmail="pBm"
message "This order needs to be held. Hugs are appreciated. The manifest has been cleared."
    if vmail CONTAINS "p"
call DeleteLastManifested
    endif
endif

if vmail contains "B"
if ShipCode contains "P" AND OrderPickedUp=""
message "This pickup is supposed to be on the dock. The manifest has been cleared."
call DeleteLastManifested
endif
endif

if vmail contains "x"
if Zip=0
message "This order is lacking its mailing zip code. The mail flag has been cleared. Please correct and try again."
call DeleteLastManifested
endif
endif

;if  info("Selected") <  info("Records") AND Order≠"" AND Order notcontains "0000"
;downrecord
   ; if PickSheet≠"" AND OriginalPicksheet=""
   ; OriginalPicksheet=PickSheet
   ; endif
;endif

    if zcancel="Yes" AND «1stTotal»>0
        NoYes "Did you issue a refund?"
        if clipboard()="No" ;because they never really had an order in the system
        «1stTotal»=0
        endif
    endif

;If ShipCode CONTAINS "X" AND vmail NOTCONTAINS "pB"
;call ".flatrate"
;Message "Should this order be shipped in a flat rate box? Check with Bernice or Lou if you're not sure."
;endif

if zcanada="Yes" AND zcancel="No"
if AdjTotal>2500
message "This order is over $2500. Please check for harmonized code issues (or ask Ryan or Bernice to do it)."
endif
;    if «3SpareNumber»≤64 and AdjTotal≤400
;message "Remember to fill out the SHORT form for the post office, assuming this is first class mail to Canada."
;    else
;message "Remember to fill out the LONG form for the post office (all priority mail and over $400 to Canada)."
;YesNo "Do you need a second copy of the picksheet?"
;            if clipboard()="Yes"
;            call ReprintPicksheet
;            endif
;    endif
endif
        zcancel="No"

if zstat=""
call ".session"
endif
    endif
zcanada="No" 
˛/.doneˇ˛.flatrateˇ;I believe this is no longer being called
zflatrate=""
if zcanada="No"
if vmail≠""

;zones 1&2
If (Zip ≥ 03800 AND Zip ≤ 04999)
Message "If this order weighs more than 10#, it should be in a flat rate box."
zflatrate="ok"
endif

;zone 3
If (Zip ≥ 01000 AND Zip ≤ 03799) OR (Zip ≥ 05000 AND Zip ≤ 06599) OR (Zip ≥ 06700 AND Zip ≤ 06799) OR 
    (Zip ≥ 12000 AND Zip ≤ 12399) OR (Zip ≥ 12800 AND Zip ≤ 12999)
Message "If this order weighs more than 7#, it should be in a flat rate box."
zflatrate="ok"
endif

;zone 4
If (Zip ≥ 00500 AND Zip ≤ 00599) OR (Zip ≥ 06600 AND Zip ≤ 06699) OR (Zip ≥ 06800 AND Zip ≤ 11999) OR 
    (Zip ≥ 12400 AND Zip ≤ 12799) OR (Zip ≥ 13000 AND Zip ≤ 14999) OR (Zip ≥ 15500 AND Zip ≤ 15599) OR 
    (Zip ≥ 15700 AND Zip ≤ 15999) OR (Zip ≥ 16300 AND Zip ≤ 22399) OR (Zip ≥ 22600 AND Zip ≤ 22699) OR 
    (Zip ≥ 25400 AND Zip ≤ 25499) OR (Zip ≥ 26700 AND Zip ≤ 26799)
Message "If this order weighs more than 6#, it should be in a flat rate box."
zflatrate="ok"
endif

;zones 5-7
If (Zip ≥ 15000 AND Zip ≤ 15499) OR (Zip ≥ 15600 AND Zip ≤ 15699) OR (Zip ≥ 16000 AND Zip ≤ 16299) OR 
    (Zip ≥ 22400 AND Zip ≤ 22599) OR (Zip ≥ 22700 AND Zip ≤ 25399) OR (Zip ≥ 25500 AND Zip ≤ 26699) OR 
    (Zip ≥ 26800 AND Zip ≤ 58899) OR (Zip ≥ 59200 AND Zip ≤ 59399) OR (Zip ≥ 60000 AND Zip ≤ 73199) OR 
    (Zip ≥ 73400 AND Zip ≤ 76799) OR (Zip ≥ 77000 AND Zip ≤ 77899) OR (Zip ≥ 79200 AND Zip ≤ 79299) OR 
    (Zip ≥ 82200 AND Zip ≤ 82299) OR (Zip ≥ 82700 AND Zip ≤ 82799)
Message "If this order weighs more than 3#, it should be in a flat rate box."
zflatrate="ok"
endif

;zone 8
If (Zip ≥ 00600 AND Zip ≤ 00999) OR (Zip ≥ 59000 AND Zip ≤ 59199) OR (Zip ≥ 59400 AND Zip ≤ 59999) OR 
    (Zip ≥ 73300 AND Zip ≤ 73399) OR (Zip ≥ 76800 AND Zip ≤ 76999) OR (Zip ≥ 77900 AND Zip ≤ 79199) OR 
    (Zip ≥ 79300 AND Zip ≤ 82199) OR (Zip ≥ 82300 AND Zip ≤ 82699) OR (Zip ≥ 82800 AND Zip ≤ 99999)
Message "If this order weighs more than 2#, it should be in a flat rate box."
zflatrate="ok"
endif

;to catch exceptions
if zflatrate=""
message "The zip for this mail-only order was not recognized. Please check with Gene. We may need to figure out what's up."
stop
endif 

zflatrate=""

endif
endif˛/.flatrateˇ˛.shippingˇwaswindow=info("windowname")
global «$Sh», VZ, Vship
VZ=Z
Vship=ShippingWt
If ShipCode Contains "U" or ShipCode Contains "X" or ShipCode Contains "H"
Openfile zcomyear+"shiplookup"
Find ZipBegin<VZ And VZ≤ZipEnd
Case Vship≤ 2
«$Sh»=«≤2»
Case Vship ≤ 5
«$Sh»=«≤5»
Case Vship≤10
«$Sh»=«≤10»
Case Vship≤15
«$Sh»=«≤15»
Case Vship≤20
«$Sh»=«≤20»
Case Vship≤25
«$Sh»=«≤25»
Case Vship≤30
«$Sh»=«≤30»
Case Vship≤35
«$Sh»=«≤35»
Case Vship≤45
«$Sh»=«≤45»
Case Vship<200
«$Sh»=Vship*«>45»
Case Vship<500
«$Sh»=Vship*«≥200»
Case Vship≥500
«$Sh»=Vship*«≥500»
EndCase
closefile
window waswindow
«$Shipping»=«$Sh»
Field Subtotal
Else
If ShipCode Contains "C" And  «$Shipping»=0
Field «$Shipping»
EditCell
Field Subtotal
Else «$Shipping» =«$Shipping»
Field Subtotal
Endif
EndIf
˛/.shippingˇ˛.shipping_canadaˇcase zshipwt ≤ 8 AND zshipwt>0
	«$Shipping»=14.15 ;first class package mail international
case zshipwt > 8 AND zshipwt ≤ 32
	«$Shipping»=21.00
case zshipwt > 32 AND zshipwt ≤ 48
	«$Shipping»=32.00
case zshipwt > 48 AND zshipwt ≤ 64
	«$Shipping»=43.15
case zshipwt > 64 AND zshipwt ≤ 144 ;medium flat rate priority mail international
	«$Shipping»=53.95
case zshipwt > 144 AND zshipwt ≤ 240 ;large flat rate priority mail international
	«$Shipping»=70.15
case zshipwt > 240 AND zshipwt ≤ 256 ;regular priority mail international
	«$Shipping»=81.75
case zshipwt > 256 AND zshipwt ≤ 272
	«$Shipping»=84.65
case zshipwt > 272 AND zshipwt ≤ 288
	«$Shipping»=87.60
case zshipwt > 288 AND zshipwt ≤ 304
	«$Shipping»=90.55
case zshipwt > 304 AND zshipwt ≤ 320
	«$Shipping»=93.55
case zshipwt > 320 AND zshipwt ≤ 336
	«$Shipping»=96.50
case zshipwt > 336 AND zshipwt ≤ 352
	«$Shipping»=99.40
case zshipwt > 352 AND zshipwt ≤ 368
	«$Shipping»=102.35
case zshipwt > 368 AND zshipwt ≤ 384
	«$Shipping»=105.30
case zshipwt > 384 AND zshipwt ≤ 400
	«$Shipping»=108.25
case zshipwt > 400 AND zshipwt ≤ 416
	«$Shipping»=111.25
case zshipwt > 416 AND zshipwt ≤ 432
	«$Shipping»=114.15
case zshipwt > 432 AND zshipwt ≤ 448
	«$Shipping»=117.10
case zshipwt > 448 AND zshipwt ≤ 464
	«$Shipping»=120.05
case zshipwt > 464 AND zshipwt ≤ 480
	«$Shipping»=123.00
case zshipwt > 480 AND zshipwt ≤ 496
	«$Shipping»=125.95
case zshipwt > 496 AND zshipwt ≤ 512
	«$Shipping»=128.90
case zshipwt > 512 AND zshipwt ≤ 528
	«$Shipping»=131.85
case zshipwt > 528 AND zshipwt ≤ 544
	«$Shipping»=134.80
case zshipwt > 544 AND zshipwt ≤ 560
	«$Shipping»=137.75
case zshipwt > 560 AND zshipwt ≤ 576
	«$Shipping»=140.70
case zshipwt > 576 AND zshipwt ≤ 592
	«$Shipping»=143.65
case zshipwt > 592 AND zshipwt ≤ 608
	«$Shipping»=146.60
case zshipwt > 608 AND zshipwt ≤ 624
	«$Shipping»=149.55
case zshipwt > 624 AND zshipwt ≤ 640
	«$Shipping»=152.50
case zshipwt > 640 AND zshipwt ≤ 656
	«$Shipping»=155.45
case zshipwt > 656 AND zshipwt ≤ 672
	«$Shipping»=158.40
case zshipwt > 672 AND zshipwt ≤ 688
	«$Shipping»=161.30
case zshipwt > 688 AND zshipwt ≤ 704
	«$Shipping»=162.30
case zshipwt > 704 AND zshipwt ≤ 720
	«$Shipping»=163.30
case zshipwt > 720 AND zshipwt ≤ 736
	«$Shipping»=164.30
case zshipwt > 736 AND zshipwt ≤ 752
	«$Shipping»=165.30
case zshipwt > 752 AND zshipwt ≤ 768
	«$Shipping»=166.30
case zshipwt > 768 AND zshipwt ≤ 784
	«$Shipping»=167.30
case zshipwt > 784 AND zshipwt ≤ 800
	«$Shipping»=168.30
case zshipwt > 800
	message "This order's shipping needs to be checked. Please check with Ryan."
endcase

if AdjTotal<30 AND AdjTotal≠0 AND («1stTotal»-«$Shipping»)<30
«$Shipping»=«$Shipping»+6
endif˛/.shipping_canadaˇ˛.sessionˇIf info("selected") < info("records")
Stop
EndIf
getscrapok "Next!"
clipboard()=clipboard()["0-9",-1]

if clipboard() = ""
    if vmail≠"BG"
vmail=""
    endif
stop
endif

    if str(clipboard()) notcontains "."
Numb=int(val(clipboard())) ;6-2 change
    else
Numb=val(clipboard())
    endif    
find OrderNo=Numb
if  info("Found") 
    case datepattern(today(),"Day") contains "Mon" AND today()-FillDate>3 AND MailFlag≠""
    MailFlag=""
    case today()-FillDate>2 AND datepattern(today(),"Day") notcontains "Mon" AND MailFlag≠""
    MailFlag=""
    endcase
    if PickSheet="" AND Order notcontains "0000" AND Order≠"" AND Order notcontains "1)"+¬+"0"+¬+¬+"0"
    Message "This picksheet's never been run!"
    stop
    endif
       if Status=""
Checker=vchecker
    If Checker=""
    call .checkers
    Endif
vchecker=Checker
field Puller
call .pullers
        endif
else
            YesNo "Nothing found (i.e. do you really know what you're doing?), try again?"
             if clipboard()="Yes"
             call .session
            endif
endif
if vmail contains "B"
call .done
endif˛/.sessionˇ˛.checkersˇ    if  info("Files") notcontains "pullers&checkers"
    openfile "pullers&checkers"
    window "Hide This Window"
    window waswindow
    endif
zchecker=""
GetScrapOK "Who's the checker?"

If clipboard()=""
YesNo "Is there a checker?"
    if clipboard()="No"
Checker=""
stop
    endif
    If clipboard()="Yes"
call .checkers   
    endif 
endif

if clipboard()≠""
    zcode=clipboard()
    zchecker=lookup("pullers&checkers","code",zcode,"checker_name","",0)
        if zchecker≠""
    Checker=zchecker
        endif    
endif

if zchecker=""
Message "You haven't assigned a valid checker, try again!"
call .checkers
endif˛/.checkersˇ˛.pullersˇ    if  info("Files") notcontains "pullers&checkers"
    openfile "pullers&checkers"
    window "Hide This Window"
    window waswindow
    endif
zpuller=""
GetScrapOK "Who's the puller?"

If clipboard()=""
YesNo "Is there a puller?"
    if clipboard()="No"
Puller=""
stop
    endif
    If clipboard()="Yes"
call .pullers   
    endif 
endif

if clipboard()≠""
    zcode=clipboard()
    zpuller=lookup("pullers&checkers","code",zcode,"puller_name","",0)
        if zpuller≠""
    Puller=zpuller
        endif    
endif

if zpuller=""
Message "You haven't assigned a valid puller, try again!"
call .pullers
endif˛/.pullersˇ˛.manifestˇdata= info("DatabaseName") 

    if  info("Files") notcontains "MailManifest"
    openfile "MailManifest"
    window waswindow
    endif

if arraysize(Order,¶)=1 AND Order CONTAINS "backorder"
goto NoManifest
endif

    if str(OrderNo) notcontains "."
window "MailManifest:packages"

arraybuild zmanifest,",","",«OrderNo»
;message zmanifest
zshipment=str(grabdata(data,OrderNo))
;message zshipment

    if zmanifest notcontains zshipment
lastrecord
InsertBelow
Contact=grabdata(data,Con)
Group=grabdata(data,Group)
«Sh Add1»=grabdata(data,MAd)
City=grabdata(data,City)
St=grabdata(data,St)
Zip=grabdata(data,Zip)
«OrderNo»=grabdata(data,OrderNo)
«UPS#»="mail"+?(zcanada="Yes","-"+grabdata(data,«9SpareText»),"")
«C#»=str(grabdata(data,«C#»))
;Code=?(vmail contains "A","A",?(vmail contains "B","B",?(vmail contains "C","C","")))
Code=?(vmail contains "x" OR vmail contains "p" OR vmail contains "u",vmail[2,2],vmail)
email=?(grabdata(data,Email)≠"",grabdata(data,Email),"")
    else
NoYes "This order is already on your mail manifest. Do you want to add another package for the same order?"
        if clipboard()="Yes"
lastrecord
InsertBelow
Contact=grabdata(data,Con)
Group=grabdata(data,Group)
«Sh Add1»=grabdata(data,MAd)
City=grabdata(data,City)
St=grabdata(data,St)
Zip=grabdata(data,Zip)
«OrderNo»=grabdata(data,OrderNo)
«UPS#»="mail"+?(zcanada="Yes","-"+grabdata(data,«9SpareText»),"")
«C#»=str(grabdata(data,«C#»))
;Code=?(vmail contains "A","A",?(vmail contains "B","B",?(vmail contains "C","C","")))
Code=?(vmail contains "x" OR vmail contains "p",vmail[2,-1],vmail)
email=?(grabdata(data,Email)≠"",grabdata(data,Email),"")
        else
Message "Nothing was added to the manifest."
        endif        
    endif

save
window waswindow
    else
    message "This is a piece of a group. You can only manifest packages from the cover sheet for a group."
    endif
    
NoManifest:
˛/.manifestˇ˛.addresschangedˇ«address_changes_date»=today()
«address_changes»=?(«address_changes»="",info("fieldname"),
                        ?(«address_changes» notcontains info("fieldname"),
                        «address_changes»+", "+info("fieldname"),«address_changes»))˛/.addresschangedˇ˛.clearaddresschangeˇ«address_changes_date»=0
«address_changes»=""˛/.clearaddresschangeˇ˛.check_member_listˇlocal zlastname, zgroupname, zmemnum, zaddressmail, zemail, zphone
waswindow=zyear+"seedstally:seedspagecheck"
    if info("files") notcontains "member_list"
    openfile "member_list"
    window waswindow
    endif

    if Con[-1,-1]=" "
zlastname=Con[1,-2]["- ",-1][2,-1]
    else
zlastname=Con["- ",-1][2,-1]
    endif
zgroupname=Group
zmemnum=«C#»
zaddressmail=MAd
zemail=Email
zphone=Telephone
;message zlastname

window "member_list"
selectall
;select Con contains zlastname
select Con contains zlastname OR Group contains zgroupname OR «C#» = zmemnum
    OR  MAd contains zaddressmail OR email contains zemail OR phone contains zphone
window waswindow˛/.check_member_listˇ˛.check_seeds_allˇlocal zlastname, zgroupname, zmemnum, zaddressmail, zemail, zphone
waswindow=zyear+"seedstally:seedspagecheck"
    if info("files") notcontains "seeds all"
    openfile "seeds all"
    window waswindow
    endif

    if Con[-1,-1]=" "
zlastname=Con[1,-2]["- ",-1][2,-1]
    else
zlastname=Con["- ",-1][2,-1]
    endif
zgroupname=Group
zmemnum=«C#»
zaddressmail=MAd
zemail=Email
zphone=Telephone
;message zlastname

openfile "seeds all"
selectall
select MAd contains zaddressmail OR Group contains zgroupname OR («C#» = zmemnum AND «C#» ≠ 0)
    OR Con contains zlastname OR Email contains zemail OR Telephone contains zphone
window waswindow˛/.check_seeds_allˇ˛.zoneˇcase Zip ≥ 3800 AND Zip ≤ 4999
zone=2

case Zip ≥ 1000 AND Zip ≤ 3799
zone=3

case Zip ≥ 5000 AND Zip ≤ 6599
zone=3

case Zip ≥ 6700 AND Zip ≤ 6799
zone=3

case Zip ≥ 12000 AND Zip ≤ 12399
zone=3

case Zip ≥ 12800 AND Zip ≤ 12999
zone=3

case Zip ≥ 500 AND Zip ≤ 599
zone=4

case Zip ≥ 6600 AND Zip ≤ 6699
zone=4

case Zip ≥ 6800 AND Zip ≤ 11999
zone=4

case Zip ≥ 12400 AND Zip ≤ 12799
zone=4

case Zip ≥ 13000 AND Zip ≤ 14999
zone=4

case Zip ≥ 15500 AND Zip ≤ 15599
zone=4

case Zip ≥ 16300 AND Zip ≤ 21299
zone=4

case Zip ≥ 21400 AND Zip ≤ 22399
zone=4

case Zip ≥ 22600 AND Zip ≤ 22699
zone=4

case Zip ≥ 25400 AND Zip ≤ 25499
zone=4

case Zip ≥ 26700 AND Zip ≤ 26799
zone=4

case Zip ≥ 15000 AND Zip ≤ 15499
zone=5

case Zip ≥ 15600 AND Zip ≤ 15699
zone=5

case Zip ≥ 16000 AND Zip ≤ 16299
zone=5

case Zip ≥ 22400 AND Zip ≤ 22599
zone=5

case Zip ≥ 22700 AND Zip ≤ 25399
zone=5

case Zip ≥ 25500 AND Zip ≤ 26699
zone=5

case Zip ≥ 26800 AND Zip ≤ 26899
zone=5

case Zip ≥ 27000 AND Zip ≤ 29799
zone=5

case Zip ≥ 37600 AND Zip ≤ 37999
zone=5

case Zip ≥ 40000 AND Zip ≤ 41899
zone=5

case Zip ≥ 42500 AND Zip ≤ 42799
zone=5

case Zip ≥ 43000 AND Zip ≤ 47499
zone=5

case Zip ≥ 47900 AND Zip ≤ 49999
zone=5

case Zip ≥ 53000 AND Zip ≤ 53299
zone=5

case Zip ≥ 53400 AND Zip ≤ 53499
zone=5

case Zip ≥ 54100 AND Zip ≤ 54399
zone=5

case Zip ≥ 54500 AND Zip ≤ 54599
zone=5

case Zip ≥ 54900 AND Zip ≤ 54999
zone=5

case Zip ≥ 60000 AND Zip ≤ 60999
zone=5

defaultcase
zone=6

endcase
˛/.zoneˇ˛nextorder/1ˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
zcancel="No"
form= info("FormName") 
If info("selected") < info("records")
SelectAll
EndIf
vmail=""
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
field OrderNo
sortup
save

    if form= "seedspagecheck"
    if OrderNo<30000 OR (OrderNo>70000 AND OrderNo<1000000) ;6-2 change
;    YesNo "Are you processing mail?"
;        if clipboard()="Yes"
        vmail="p"
            NoYes "Are you processing backorders?"
            if clipboard()="Yes"
            vmail=vmail+"B"
                NoYes "Are you tallying backorders that have no seeds available that are not all backorders?"
                if clipboard()="Yes"
                vmail=vmail+"m"
                endif
                NoYes "Are you running a batch of group pieces?"
                if clipboard()="Yes"
                vmail=vmail+"g"
                endif
            endif
;        endif
;        if vmail=""
;        vmail="u"
;            NoYes "Are you processing backorders?"
;            if clipboard()="Yes"
;            vmail=vmail+"B"
;                NoYes "Are you tallying backorders that have no seeds available that are not all backorders?"
;                if clipboard()="Yes"
;                vmail=vmail+"m"
;                endif
;            endif
;        endif
    else
    message "You're on the wrong form to be running these orders."
    stop
    endif    
    endif
;message vmail    
getscrapok "What's the order number?"
clipboard()=clipboard()["0-9",-1]

if clipboard() = ""
vmail=""
stop
endif

    if str(clipboard()) notcontains "."
Numb=int(val(clipboard())) ;6-2 change
    else
Numb=val(clipboard())
    endif    
find OrderNo=Numb
        if  info("Found")
    case datepattern(today(),"Day") contains "Mon" AND today()-FillDate>3 AND MailFlag≠""
    MailFlag=""
    case today()-FillDate>2 AND datepattern(today(),"Day") notcontains "Mon" AND MailFlag≠""
    MailFlag=""
    endcase
    if PickSheet="" AND Order notcontains "0000" AND Order≠"" AND Order notcontains "1)"+¬+"0"+¬+¬+"0"
    Message "This picksheet's never been run!"
    stop
    endif
                if Status=""
                If Checker=""
                Call .checkers
                endif
        vchecker=Checker
        field Puller
        call .pullers
                endif
        else
            YesNo "Nothing found (i.e. do you really know what you're doing?), try again?"
             if clipboard()="Yes"
             call .session
            endif
        endif
    endif
 If (Order="" OR Order contains "0000" OR Order contains "1)"+¬+"0"+¬+¬+"0")AND AdjTotal>0 AND Status=""
 downrecord
 call .pullers
 endif   
 if vmail contains "B"
call .done
endif˛/nextorder/1ˇ˛setchecker/2ˇcall .checkers
vchecker=Checker˛/setchecker/2ˇ˛ChangePuller/3ˇcall .pullers˛/ChangePuller/3ˇ˛justlocate/4ˇzcancel="No"
form= info("FormName") 
If info("selected") < info("records")
SelectAll
EndIf
if folderpath(dbinfo("folder","")) NOTCONTAINS "Facilitator" /*added by Rachel/Ken to stop a call to vmail error in facilitator files*/
    if vmail≠"BG"
    vmail=""
    endif
    endif 
getscrap "What's the order number?"
clipboard()=clipboard()["0-9",-1]
    if str(clipboard()) notcontains "."
Numb=int(val(clipboard())) ;6-2 change
    else
Numb=val(clipboard())
    endif    
find OrderNo=Numb OR «5SpareNumber»=Numb
if info("Found")
    case datepattern(today(),"Day") contains "Mon" AND  (today()-4 ≥ «additional_mail») AND MailFlag≠"" AND FillDate≠today()
    MailFlag=""
    case  (today()-3 ≥ «additional_mail») AND datepattern(today(),"Day") notcontains "Mon" AND MailFlag≠"" AND FillDate≠today()
    MailFlag=""
    endcase
stop
else
YesNo "Nothing found, try again?"
    if clipboard()="Yes"
    call justlocate/4
    endif
endif˛/justlocate/4ˇ˛multipagetotal/5ˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
rayg="M"
ODisc=0
        if vmail="BG"
            getscrap "Order to consolidate"
        else
        YesNo "Is "+str(int(Numb))+" the group order you want to consolidate?"
            if clipboard()="Yes"
            clipboard()=str(int(Numb))
            else
            getscrap "Order to consolidate"
            endif
        endif
;    if clipboard()="Yes"
;    clipboard()=str(Numb)[1,5]
;    else
;    getscrap "Order to consolidate"
;    endif
Numb=val(clipboard())
Select int(OrderNo)=Numb
    if info("Selected") =1 OR (info("Selected") = info("Records"))
  message "Whoa!!! This is not a group order. Check it owwwt!"
    stop
   endif

Field OrderNo
SortUp
FirstRecord
    case datepattern(today(),"Day") contains "Mon" AND today()-FillDate>3 AND MailFlag≠""
    MailFlag=""
    case today()-FillDate>2 AND datepattern(today(),"Day") notcontains "Mon" AND MailFlag≠""
    MailFlag=""
    endcase

            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
    if Status=""
    Checker=vchecker
        If Checker=""
        call .checkers
        Endif
    vchecker=Checker
    if Puller=""
call .pullers
    endif
    endif
ShippingWt=0
TaxTotal=0
Subtotal=0
AddOns=0
    if zcanada="Yes"
«1SpareNumber»=0
«2SpareNumber»=0
    endif

;David and I agreed to not cover the unlikely case that the entire group does not get the member discount 12-26-12
;if MemDisc<.01*AdjTotal
;MemDisc=0
;Field MemDisc
;Total
;endif

Field Subtotal
Total
;Field TaxTotal
;Total
Field ShippingWt
Total
Field AddOns
Total
Field OGSTotal
Total
Field «1SpareNumber»
Total
Field «2SpareNumber»
Total
LastRecord
sub=Subtotal
;tax=TaxTotal
Vship=ShippingWt
    if zcanada="Yes"
zseedwt=«1SpareNumber»
zpacketwt=«2SpareNumber»
    endif
ODisc=OGSTotal
atotal=AddOns
RemoveSummaries 7
FirstRecord
Subtotal=sub
;moved sales tax accounting to .salestax
OGSTotal=ODisc
ShippingWt=Vship
    if zcanada="Yes"
«1SpareNumber»=zseedwt
«2SpareNumber»=zpacketwt
    endif

;    if MemDisc=0
;    MemDisc=ODisc
;    vmemflag=1 ;this needs confirming with office protocol - see .retotal - when not set to zero it does not recalculate MemDisc on coversheet
;    endif

AddOns=atotal
call ".retotal"
arrayselectedbuild members,¶,zyear+"seedstally",?(OrderNo>Numb,str(round((OrderNo - Numb)*1000,1))
            +". "+Con+" - $"+str(Subtotal+AddOns),"")
Field Puller
    endif
zcanada="No"
˛/multipagetotal/5ˇ˛reworkchanges/®ˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
waswindow= info("WindowName") 
form= info("FormName") 
zcanada="No"
if PickSheet=""
Message "This works only after the picksheet has been run. Ask Ryan if you need to adjust
                this order anyway."
stop
endif

Case form="seedspagecheck"
if OrderNo <30000 OR (OrderNo>70000 AND OrderNo<1000000)

if zbackrunning="No"
    openfile zcomyear+"SeedsComments linked"
    save
    closefile
    openfile zcomyear+"SeedsComments"
    openfile "&&"+zcomyear+"SeedsComments linked"
    save
    window "Hide This Window"
    window waswindow
endif    
    
if addon≠""
            if Zip = 0 and «9SpareText» ≠ ""
            message "This is a Canadian order. We don't allow add-ons. Talk to Ryan or Bernice. This macro will stop"
            stop
            endif
rayj=""
linenu=1
loop
rayj=extract(addon,¶,linenu)
stoploopif rayj=""
Order=Order+¶+¬+extract(rayj,chr(43),1)+¬+extract(rayj,chr(43),2)+¬+¬+¬+extract(rayj,chr(43),3)
linenu=linenu+1
until  info("Empty") 
endif

vmem=0
if MemDisc>0 OR Discount=.5
vmem=1
endif
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
sub=Subs
    openfile "SeedsTotaller"
        oldwindow=info("windowname")
        window waswindow
        order=""
        linenu=1
        loop
        order=order+extract(extract(Order,¶,linenu),¬,1)+¬+?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[1,"-"][1,-2],extract(extract(Order,¶,linenu),¬,2))+¬+
          ?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[-1,-1],"")+¬+extract(extract(Order,¶,linenu),¬,3)+¬+
          extract(extract(Order,¶,linenu),¬,4)+¬+extract(extract(Order,¶,linenu),¬,5)+¬+
          extract(extract(Order,¶,linenu),¬,6)+¬+strip(extract(extract(Order,¶,linenu),¬,7))+¬+¬+
          extract(extract(Order,¶,linenu),¬,8)+¶
        linenu=linenu+1
        Until  extract(extract(Order,¶,linenu),¬,2)=""   
        window "SeedsTotaller"
else
message "You're on the wrong form to be running these orders."
stop
endif
endcase
    
openfile "&@order"
;stop
call ".recalculate"
    if Order contains "backorder"
    if brange=""
    call ".backorder"
    endif
    endif
        if addon≠""
        «BalDue/Refund»=Paid-GrTotal
        message "Please remember to press BACKORDER or COMPLETE to fully update this order's records when you're done."
        endif
addon=""
zcanada="No"
    endif
˛/reworkchanges/®ˇ˛additionalpay/åˇif folderpath(dbinfo("folder","")) CONTAINS "tally.seeds" OR folderpath(dbinfo("folder","")) CONTAINS "facilitator"
zmethod=""
    if  info("Selected") < info("Records") 
    selectall
    endif
    
    if  info("Windows") notcontains "additional_payments"
    WindowBox "30 640 300 1025"
    openform "additional_payments"
    else
    goform "additional_payments"
    zoomwindow 30,640,300,500,""
    endif
    
;message "See Gene to take care of any additional payments or payment adjustments. Thank you."
stop


;new macro from Alice 7/1/20. Seems we do not need to use it this way and can keep our previous macro with the "sidecar". As long as adding a "transfer" choice will work out.
; local addpay, vhow,vorderref,vorderpay
 
 ;GetScrap  "What's the additional payment?" ;dollar amount
 ;addpay=val(clipboard())
 ;getscrap "How was it paid?"
 ;case ((clipboard() contains "cc" or clipboard() contains "cred") and (today()-OrderPlaced≥90 or (OrderNo≥700001 and OrderNo<710000)))
;Message "Can't do a CC Refund: Issue a Check"
;stop
; case (clipboard() contains "cc" or clipboard() contains "credit")
; vhow="Credit_Card"
 ;case (clipboard() contains "gc" or clipboard() contains "gift")
 ;vhow="Gift_Certificate"
 ;case (clipboard() contains "ch" or clipboard() contains "√")
 ;vhow="Check"
 ;case (clipboard() contains "cash" or clipboard() contains "$")
 ;vhow="Cash"
 ;endcase
 ;case clipboard() contains "tr"
 ;vhow="Transfer"
 ;vorderref=str(OrderNo)
 ;gettext "To which order are you transferring this payment?",vorderpay
 ;«Notes2»=«Notes2»+¶+"Payment transferred to order "+vorderpay
 ;endcase
 
; Paid= Paid+addpay
 
 ;«BalDue/Refund»=«BalDue/Refund»+addpay
 ;If «AddPay1»=0
  ;  «AddPay1»= addpay
   ; «DatePay1»=today()
    ;«MethodPay1»=vhow
;else
 ;   if «AddPay2»=0
  ;      «AddPay2»= addpay
   ;     «DatePay2»=today()
    ;    «MethodPay2»=vhow
    ;else
     ;   if «AddPay3»=0
      ;      «AddPay3»=addpay
       ;     «DatePay3»=today()
        ;    «MethodPay3»=vhow
    ;    else
     ;       if «AddPay4»=0
      ;      «AddPay4»=addpay
       ;     «DatePay4»=today()
        ;    «MethodPay4»=vhow
         ;   else
          ;      if «AddPay5»=0
           ;     «AddPay5»=addpay
            ;    «DatePay5»=today()
             ;   «MethodPay5»=vhow
              ;  else
               ;     if «AddPay6»=0
                ;    «AddPay6»=addpay
                 ;   «DatePay6»=today()
                  ;  «MethodPay6»=vhow
                  ;  else 
                  ;  message "This order is crazy. Please consolidate additional payments."
                   ; stop
                   ; endif
             ;   endif
          ; endif
       ; endif
   ; endif
endif

;if vhow="Transfer"
;find OrderNo=val(vorderpay)
;    if info("found")
 ;   addpay=-addpay
  ;  «Notes2»=«Notes2»+¶+"Payment transferred from order "+str(vorderref)
   ;  Paid= Paid+addpay
 
; «BalDue/Refund»=«BalDue/Refund»+addpay
; If «AddPay1»=0
 ;   «AddPay1»= addpay
  ;  «DatePay1»=today()
   ; «MethodPay1»=vhow
;else
 ;   if «AddPay2»=0
  ;      «AddPay2»= addpay
   ;     «DatePay2»=today()
    ;    «MethodPay2»=vhow
  ;  else
   ;     if «AddPay3»=0
    ;        «AddPay3»=addpay
     ;       «DatePay3»=today()
      ;      «MethodPay3»=vhow
     ;   else
      ;      if «AddPay4»=0
       ;     «AddPay4»=addpay
        ;    «DatePay4»=today()
         ;   «MethodPay4»=vhow
          ;  else
          ;      if «AddPay5»=0
           ;     «AddPay5»=addpay
            ;    «DatePay5»=today()
             ;   «MethodPay5»=vhow
              ;  else
               ;     if «AddPay6»=0
                ;    «AddPay6»=addpay
        ;            «DatePay6»=today()
         ;           «MethodPay6»=vhow
          ;          else 
           ;         message "This order is crazy. Please consolidate additional payments."
            ;        stop
             ;       endif
            ;    endif
   ;        endif
    ;    endif
  ;  endif
;endif
;yesno
;"Does this look right?"
;if clipboard() contains "n"
;stop
;endif
;find OrderNo=val(vorderref)
;else
;message "Can't find the order you're using for this transfer"
;endif
;endif
    
˛/additionalpay/åˇ˛commentfile/œˇ    openfile zcomyear+"SeedsComments linked"
    Synchronize
    Window zcomyear+"SeedsComments linked:linechecker_old"    
    field Item
    sortup
    firstrecord˛/commentfile/œˇ˛(details)ˇ˛/(details)ˇ˛ChangePaidˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
getscrap "New Paid"
Paid=val(clipboard())
«BalDue/Refund»=Paid-GrTotal
if Paid≠«1stPayment»+AddPay1+AddPay2+AddPay3+AddPay4+AddPay5+AddPay6
message "The amount you entered doesn't match the payment(s) recorded. Please be sure this is correct. Talk to me if you're unsure, it can avoid bookkeeping headaches. Thanks. Ryan"
endif
    endif˛/ChangePaidˇ˛AdjustDonatedRefundˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
getscrap "How much of their refund should go to Nibezun?"
Donated_Refund=val(clipboard())
getscrap "What should paid be set to before you re-Complete the order?"
Paid=val(clipboard())
«BalDue/Refund»=Paid-GrTotal
    endif˛/AdjustDonatedRefundˇ˛PicksheetPrint/πˇ;this prints all but Canadian orders, they have their own loop
oldfile=  info("DatabaseName") 
waswindow=info("windowname")
form=info("FormName")
rayg=""
zlarge=0
Synchronize
ztime=now() ;resets sync timer
field OrderNo
SortUp
select Order contains "0000" OR Order = "" OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderComments CONTAINS "empty order")
    OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderNo>1000000) ;6-2 change
    
selectreverse
selectwithin ShipCode≠"D"
selectwithin «9SpareText»="" ;exclude Canadians
;selectwithin TaxState≠""
;selectwithin (ztaxstates contains TaxState AND TaxRate≠0) 
;    OR (ztaxstates notcontains TaxState AND TaxRate=0) OR Taxable contains "N"
;    OR str(OrderNo) contains "."
    
getscrap "first order to print"
Numb=val(clipboard())
ono=Numb
pickno=Numb
Selectwithin int(OrderNo)≥Numb ;6-2 change
getscrap "last order to print"
Numb=val(clipboard())
selectwithin int(OrderNo)≤Numb ;6-2 change
if info("Empty")
message "It looks like your last order to print was lower than your first order to print. Please try again."
stop
endif


if form= "seedspagecheck" 
if OrderNo < 30000 OR (OrderNo>70000 AND OrderNo<1000000) ;6-2 change
selectwithin PickSheet=""
else
message "You're on the wrong form to be running these orders."
stop
endif        
endif
    if  info("Empty")
    message "Picksheets have already been printed for all the orders you selected."
    selectall
    stop
    endif
    
        if form="seedspagecheck"    
        openfile zcomyear+"SeedsComments linked"
        ReSynchronize
        save
        closefile
        openfile zcomyear+"SeedsComments"
        openfile "&&"+zcomyear+"SeedsComments linked"
        endif
save
window "Hide This Window"
window waswindow
call ".comments"
selectwithin PickSheet≠"" ;this works primarily with the process in the totaller to avoid orders with items marked -avoid-.
if PickSheet≠""
firstrecord
        PrintUsingForm "", "seedspicksheet"
        print ""
      Selectwithin PickSheet CONTAINS "41)"
    If  info("Empty")
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet2"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "96)"
    If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet3"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "151)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet4"
        print ""
    endif
    skip:
    
SelectAll
Select int(OrderNo)≥pickno ;6-2 change
selectwithin int(OrderNo)≤Numb ;6-2 change
zselect= info("Selected")   
        selectwithin OrderComments≠"" OR PickSheet CONTAINS "201)" OR Notes1≠""
if info("Selected")<zselect
        if zlarge>1
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!
Also, one or more order(s) has more than 200 items. Check with Bernice or Carol about this!"
       endif
        if zlarge=0
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!"
       endif
 else
 selectall
 endif    
 else
 message "Nothing to print."
 endif  ˛/PicksheetPrint/πˇ˛ReprintPicksheetˇwaswindow= info("WindowName")
if PickSheet≠""
zcanada="No"
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif

If zcanada="No"
PrintUsingForm "", "seedspicksheet"
PrintOneRecord
if PickSheet CONTAINS "41)"
PrintUsingForm "", "seedspicksheet2"
PrintOneRecord
endif
if PickSheet CONTAINS "96)"
PrintUsingForm "", "seedspicksheet3"
PrintOneRecord
endif
if PickSheet CONTAINS "151)"
PrintUsingForm "", "seedspicksheet4"
PrintOneRecord
endif

if PickSheet CONTAINS "203)"        
message "This routine only reprints the first 4 pages. You may need to reprint any remaining pages manually."
endif      
endif

If zcanada="Yes"
            message "Remember to load Canadian picksheets."
PrintUsingForm "", "canada_picksheet"
PrintOneRecord
if PickSheet CONTAINS "39)"
PrintUsingForm "", "canada_picksheet2"
PrintOneRecord
endif
if PickSheet CONTAINS "83)"
PrintUsingForm "", "canada_picksheet3"
PrintOneRecord
endif
if PickSheet CONTAINS "127)"
PrintUsingForm "", "canada_picksheet4"
PrintOneRecord
endif
if PickSheet CONTAINS "171)"
PrintUsingForm "", "canada_picksheet5"
PrintOneRecord
endif

if PickSheet CONTAINS "201)"        
message "This routine only reprints the first 4 or 5 pages. You may need to reprint any remaining pages manually."
endif      
endif

zcanada="No"
else
message "This order has never had a picksheet printed."
endif˛/ReprintPicksheetˇ˛BatchReprintPicksheetsˇYesNo "Do you really want to reprint the picksheets for all "+str(info("Selected"))+" orders that are selected?"
if clipboard()="Yes"
firstrecord
call ReprintPicksheet
    loop
    downrecord
    call ReprintPicksheet
    until  info("EOF") 
endif˛/BatchReprintPicksheetsˇ˛PrePrintAddonˇlocal prenum, presize, precount
Message "This allows the addition of items to an order before the picksheet has been run. The math will be refigured when the picksheet is run."
            if Zip = 0 and «9SpareText» ≠ ""
            message "This is a Canadian order. We don't allow add-ons. Talk to Ryan or Bernice. This macro will stop"
            stop
            endif
if PickSheet≠""
message "This only works before the picksheet has been run."
stop
endif
    loop
prenum=""
presize=""
precount=0
GetScrap "What's the item?"
prenum=clipboard()
GetScrap "What size?"
presize=upper(clipboard())
GetScrap "How many"
precount=clipboard()
    if val(prenum)=0 or val(presize)>0 or val(precount)=0
    message "You need to fill in something for each of the item, size and count parameters. Please try again."
    stop
    endif
message prenum+presize+"-"+precount
if prenum≠"" and presize≠"" and precount≠""
Order=Order+¶+¬+prenum+¬+presize+¬+precount+¬+¬+"preadd"
endif
    until prenum=""
˛/PrePrintAddonˇ˛ChangeShipCodeˇGetScrapOK "What's the new ship code?"
if clipboard()≠""
ShipCode=upper(clipboard())
endif
if ShipCode≠"U" and ShipCode≠"P" and ShipCode≠"X" and ShipCode≠"H" and ShipCode≠"D"
message "Ahoy!! What'd ya think yer dewin??? '"+ShipCode+"'"+" is not a standard code. Please correct if necessary."
endif˛/ChangeShipCodeˇ˛textport/†ˇwaswindow= info("WindowName") 
selectall
Synchronize
ztime=now() ;resets sync timer
selectwithin Z≠0 ;in theory this eliminates Canadian orders as well
selectwithin ShipCode="U"
selectadditional ShipCode="P" and OrderPickedUp="X"
openfile "Texport"
openfile "&&"+zyear+"seedstally"
call "textport"
selectall
˛/textport/†ˇ˛AddCommentˇzcomment=OrderComments
GetScrapOK "Whazzup?"
if OrderComments=""
OrderComments=clipboard()
else
OrderComments=zcomment+" "+clipboard()
endif˛/AddCommentˇ˛SelectItem/ßˇselectall
local itemnu
itemnu=""
NoYes "Search unprinted orders?"

If clipboard()="No"
GetScrapOK "what item (ex: 204-A)?"
    if clipboard() contains "-"
        if length(clipboard()[1,"-"][1,-2])=3
        itemnu=" "+clipboard()
        else
        itemnu=clipboard()
        endif
    endif
    if clipboard() notcontains "-"
        if length(clipboard())=3
        itemnu=" "+clipboard()
        else
        itemnu=clipboard()
        endif
    endif
select Order contains itemnu AND PickSheet≠""
if  info("Selected") =  info("Records") 
message "nada"
endif
endif

If clipboard()="Yes"
GetScrapOK "what item (ex: 204-A)?"
    if clipboard() CONTAINS "-"
        itemnu=clipboard()
        select (Order contains ¬+itemnu[1,"-"][1,-2]+¬+itemnu["-",-1][2,-1] OR
        Order contains ¬+"0"+itemnu[1,"-"][1,-2]+¬+itemnu["-",-1][2,-1]) AND PickSheet=""
    else
        itemnu=clipboard()
        select (Order contains ¬+itemnu+¬ OR Order contains ¬+"0"+itemnu+¬) AND PickSheet=""
    endif
if  info("Selected") =  info("Records") 
message "nada"
endif
endif˛/SelectItem/ßˇ˛.WebLineCountˇ;this macro needs fixing to stop the macro if the open file dialog is cancelled
;I'm not sure what this was for anymore 12-27-15
;I think I was trying to learn how many characters panorama was having trouble dropping in after processing
stop
OpenFile Dialog
Field «7SpareNumber»
FormulaFill arraysize(Order,¶)
Save
CloseFile˛/.WebLineCountˇ˛.ShipLookupˇNumb=int(OrderNo)

case ShipCode="U" OR ShipCode="X"
OpenFile zyear+"shipping"
selectall
Synchronize
select «OrderNo»=Numb OR «C#»= str(Numb)

case ShipCode="P" OR ShipCode="T"
YesNo "This order is a pickup. Do you still want to check shipping?"
    if clipboard()="Yes"
    OpenFile zyear+"shipping"
    selectall
Synchronize
    select «OrderNo»=Numb OR «C#»= str(Numb)
    else
    stop
    endif

case ShipCode="H"
YesNo "This order is being held for some reason. Do you still want to check shipping?"
    if clipboard()="Yes"
    OpenFile zyear+"shipping"
    selectall
Synchronize
    select «OrderNo»=Numb OR «C#»= str(Numb)
    else
    stop
    endif
endcase
if   info("Selected") < info("Records") 
stop
else
message "nothing found"
endif
˛/.ShipLookupˇ˛.emailˇif Email≠""
clipboard()="'"+Con+"' <"+Email+">"
;applescript |||
;tell application "Finder"
;	activate
;	open location "https://webmail.hostway.com/appsuite/#!&app=io.ox/mail&folder=default0/INBOX"
;end tell
;|||
else
message "There's no email address. Call or write."
endif˛/.emailˇ˛CancelOrderˇzcancel="No"
    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
NoYes "Are you sure you want to cancel this entire order?"
if clipboard()="Yes"
zcancel="Yes"
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
Checker="Bernice"
Puller="Everybody"
Order=""
PickSheet=""
Backorder=""
ShippingWt=0
TaxTotal=0
Subtotal=0
SalesTax=0
AdjTotal=0
VolDisc=0
MemDisc=0
Surcharge=0
«$Shipping»=0
AddOns=0
OrderTotal=0
Donation=0
Membership=0
GrTotal=0
RealTax=0
Patronage=0
    if zcanada="Yes"
«1SpareNumber»=0
«2SpareNumber»=0
«3SpareNumber»=0
«4SpareNumber»=0
«3SpareMoney»=0
    endif
call .retotal
Notes1=Notes1+" "+"Order cancelled "+datepattern(today(),"mm/dd/yy")+"."
call .done
endif
    endif
zcancel="No"
zcanada="No"
 ˛/CancelOrderˇ˛.addontaxˇ;remember to adjust this to not add RealTax to group pieces
            if Zip = 0 and «9SpareText» ≠ ""
            message "This is a Canadian order. We don't allow add-ons. Talk to Ryan or Bernice. This macro will stop"
            AddOns=0
            stop
            endif
tax=0
if Taxable="Y"
GetScrap "How much sales tax is included?"
tax=clipboard()
YesNo "$"+Pattern(val(tax),"#.##")+" was paid as sales tax?"
    If clipboard()="Yes"
 
        if AddOnTax=0
AddOnTax=val(tax)
RealTax=SalesTax+AddOnTax
        else

        if AddOnTax≠0
        YesNo "Is this in addition to the tax already recorded?"
        if clipboard()="Yes"
AddOnTax=AddOnTax+val(tax)
RealTax=SalesTax+AddOnTax
        endif
       if clipboard()="No"
AddOnTax=val(tax)
RealTax=SalesTax+AddOnTax
        endif
        endif
        endif

    endif
endif
call .retotal
message "Remember to hit Complete or Backorder when you're done, please."
˛/.addontaxˇ˛(mail)ˇ˛/(mail)ˇ˛Xcountˇselectall
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
select MailFlag contains "x"
if info("Selected")<info("Records")
message "There are "+str(info("Selected"))+" mailers ready to print."
else
message "There's no mail ready to print."
endif
selectall˛/Xcountˇ˛PrintAndManifest X'sˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
form= info("FormName") 
If info("selected") < info("records")
SelectAll
EndIf
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup

message "Remember to pull out the pick sheet tray on the printer."

select MailFlag CONTAINS "x"
If info("selected") < info("records")
message "Remember to feed the label sheets one at a time."
GoForm "warehouse mail small"
print ""
else
    if info("Selected")= info("Records") 
    message "There's nothing to mail."
    endif
endif

GoForm form
    endif

firstrecord
vmail=MailFlag
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
call .manifest
            zcanada="No"
if  info("Selected") >1
loop
downrecord
vmail=MailFlag
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
call .manifest
            zcanada="No"
until info("EOF") 
endif

field MailFlag
fill ""˛/PrintAndManifest X'sˇ˛ProcessManifestˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
AlertOkCancel "For accurate shipping dates, this needs to be run on the day the mail leaves the shop. Remember, this procedure needs to be run on each computer doing paperwork."
If  info("DialogTrigger") ="OK"
window "MailManifest:packages"
selectall
field «OrderNo»
sortup
if  info("Records") >1
    firstrecord
    if «OrderNo»=0
    deleterecord
    endif
field Date
emptyfill datepattern(today(),"mm/dd/yy")

openfile zyear+"shipping"
openfile "++MailManifest"
save
closefile

vmailbatch=""
;select Date=today() and Code="" and email≠""
arraybuild vmailbatch,¶,"", ?(Date=today() and Code="" and email≠"", Contact+","+email+","+str(«OrderNo»),"")
;message vmailbatch
;clipboard()=vmailbatch
;stop

;selectall
saveacopyas "mail batch "+datepattern(today(),"mm-dd-yy")
deleteall
save

    if vmailbatch≠""
clipboard()=vmailbatch
applescript |||
tell application "Finder"
	activate
	open location "http://www.fedcoseeds.com/manage_site/send-mail-emails.php"
end tell
|||
    endif

else
message "There's nothing here to ship!"
endif
endif
    endif

window waswindow
 ˛/ProcessManifestˇ˛ManifestThisOrderˇif vmail=""
GetScrapOK "Mail flag? (A +/or B +/or C or nothing)"
clipboard()=upper(clipboard())
if clipboard()≠"A" AND clipboard()≠"B" AND clipboard()≠"C" AND clipboard()≠""
    and  clipboard() ≠ "AB" and  clipboard() ≠ "AC" and  clipboard() ≠ "BC" and  clipboard() ≠ "ABC"
message "The mail flag you entered is not valid."
stop
endif
vmail=clipboard()
endif
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
call .manifest
            zcanada="No"
˛/ManifestThisOrderˇ˛DeleteLastManifestedˇNumb=OrderNo
window "MailManifest:packages"
    if «OrderNo»=Numb
lastrecord
deleterecord
save
    else
    message "The package you attempted to delete does not match the order you're on. Nothing was deleted."
    endif
window waswindow
˛/DeleteLastManifestedˇ˛export_mail/µˇwaswindow= info("WindowName") 
;selectall
Synchronize
FirstRecord
ztime=now() ;resets sync timer
;select str(OrderNo) NOTCONTAINS "." AND Status=""
    if datepattern(today(),"Day") contains "Mon"
select Status="B/O" AND (FillDate ≥ today()-3)
selectadditional (Status≠"" AND (FillDate ≥ today()-3)) 
selectadditional («additional_mail» ≥ today()-3)
    else
select Status="B/O" AND (FillDate ≥ today()-2)
selectadditional (Status≠"" AND (FillDate ≥ today()-2)) 
selectadditional («additional_mail» ≥ today()-2)
    endif
selectwithin str(OrderNo) NOTCONTAINS "."
;selectadditional (OrderNo>88000 And OrderNo<99000)
;selectadditional OrderNo>OrderNo-300 and OrderNo>20000
selectwithin Zip≠0 OR «9SpareText»≠"" ;in theory this leaves Canadian orders in
selectwithin ShipCode="U" OR ShipCode="X" OR (ShipCode="P" and OrderPickedUp="X")
;selectadditional ShipCode="P" and OrderPickedUp="X"
openfile "Texport-mail"
openfile "&&"+zyear+"seedstally"
call "textport-mail"
window waswindow
selectall
˛/export_mail/µˇ˛export_batch_to_mailˇwaswindow= info("WindowName") 
openfile "Texport-mail"
openfile "&&"+zyear+"seedstally"
call "textport-mail"
window waswindow
;selectall
˛/export_batch_to_mailˇ˛.mailflagˇMailFlag=MailFlag[1,1]+Upper(MailFlag[2,-1])
if MailFlag ≠ "x" and MailFlag ≠ "xA" and MailFlag ≠ "xB" and MailFlag ≠ "xC" and MailFlag ≠ "" and MailFlag≠"√"
    and  MailFlag ≠ "xAB" and  MailFlag ≠ "xAC" and  MailFlag ≠ "xBC" and  MailFlag ≠ "xABC"
message "This field is only intended to flag orders for printing mailing labels. Any other use will be ignored and the flag may be deleted."
MailFlag=""
endif˛/.mailflagˇ˛.add_to_endiciaˇif vmail notcontains "B"
GetScrapOK "Mail flag? (A +/or B +/or C or nothing)"
clipboard()=upper(clipboard())
if clipboard()≠"A" AND clipboard()≠"B" AND clipboard()≠"C" AND clipboard()≠""
    and  clipboard() ≠ "AB" and  clipboard() ≠ "AC" and  clipboard() ≠ "BC" and  clipboard() ≠ "ABC"
message "The mail flag you entered is not valid."
stop
endif
MailFlag=clipboard()
«additional_mail»=today()
else
MailFlag="B"
endif˛/.add_to_endiciaˇ˛-ˇ˛/-ˇ˛(backorders)ˇ˛/(backorders)ˇ˛BackordersRun/∫ˇzitem=""
brange=""
waswindow=info("WindowName") 
form= info("FormName")
selectall
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup
firstrecord

    openfile zcomyear+"SeedsComments linked"
    synchronize 
    save
    closefile
    openfile zcomyear+"SeedsComments"
    openfile "&&"+zcomyear+"SeedsComments linked"
    save
    window "Hide This Window"
    window waswindow
    
NoYes "Did you forget to adjust Comments to account for items available only to backorders?"
    if clipboard()="Yes"
    stop
    endif
YesNo "Do you need to limit the range of orders?"
    If clipboard()="Yes"
    brange="Yes"
    else
    brange="No"
    endif
YesNo "Do you want to run only completed backorders?"
    if clipboard()="Yes"
    YesNo "Do you want to use the pre-built list of unavailable items?"
    
                if clipboard()="Yes"
                Call "backorderlist"
                endif
                
         If clipboard()="No"       
        GetScrap "Enter first item #"
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        Select «Backorder» matchexact "*"+zitem+"*"
        Loop
        GetScrap "Enter next item # or zero"
        stoploopif val(clipboard())=0
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        SelectAdditional «Backorder» matchexact "*"+zitem+"*"
        Until  val(clipboard())=0
        endif
        
        SelectReverse
        selectwithin Status="B/O"
        selectwithin Order≠""

    ;this could add items to a selection
    ;selectadditional (Backorder CONTAINS "3837" OR Backorder CONTAINS "2447" OR Backorder CONTAINS "2491") AND str(OrderNo) CONTAINS "."

        if  info("Empty") 
        message "Nothing in this set has a Status of backorder."
        selectall
        stop
        endif
            if brange="Yes"
            getscrap "first order to print"
            Numb=val(clipboard())
            Selectwithin int(OrderNo)≥Numb ;6-2 change
            getscrap "last order to print"
            Numb=val(clipboard())
            selectwithin int(OrderNo)≤Numb ;6-2 change
            endif
    else
YesNo "Do you want to print orders containing specific backordered items (things needing immediate attention)? Yes, lets you enter a list of the items."
    if clipboard()="Yes"        
        GetScrap "Enter first item #"
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        Select «Backorder» matchexact "*"+zitem+"*"
        Loop
        GetScrap "Enter next item # or zero"
        stoploopif val(clipboard())=0
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        SelectAdditional «Backorder» matchexact "*"+zitem+"*"
        Until  val(clipboard())=0
        selectwithin Status="B/O"
        selectwithin Order≠""
  
       ; this would allow us to skip group pieces or vice versa
       NoYes "Do you want to choose to run only individual or only group orders?"
       if clipboard()="Yes"
         NoYes "Do you want to select for individual orders only?"
         if clipboard()="Yes"
         selectwithin str(OrderNo) notcontains "."
         else
         NoYes "Do you want to select for group pieces only?"
         if clipboard()="Yes"
         selectwithin str(OrderNo) contains "."
         endif
         endif
       endif  
        
        if  info("Empty") 
        message "Nothing in this set has a Status of backorder."
        selectall
        stop
        endif
            if brange="Yes"
            getscrap "first order to print"
            Numb=val(clipboard())
            Selectwithin (OrderNo)≥Numb ;6-2 change
            getscrap "last order to print"
            Numb=val(clipboard())
            selectwithin int(OrderNo)≤Numb ;6-2 change
            endif
    else
        stop
    endif
    endif
    
selectwithin Backorder CONTAINS ¬+"  0"+¬ 
zbackrunning="Yes"    
FirstRecord
Loop 
    If Backorder NOTCONTAINS ¬+"  0"+¬  ;this seems unnecessary-see above
    message "Order "+str(OrderNo)+"'s backorder was recently run and either wasn't processed or failed to update the server properly. It will not be run again. Check it out!"
    «10SpareText»="ran"
    endif
If Backorder CONTAINS ¬+"  0"+¬
Call ".backordercomments"
Call "reworkchanges/®"
endif
DownRecord
Until  info("Stopped") 
    SelectWithin «10SpareText»≠"ran"  ;this seems unnecessary-see above
    openform "BackorderPicksheet"
    print ""
    CloseWindow
if arraysize(Backorder,¶)>35
message "This order has more than 35 lines of backorder. Please check with Ryan. They may not get all their backorders."
endif        
        done:
window waswindow
    YesNo "Do you want to export this set to generate a commodity list?"
    if clipboard()="Yes"
    call "exportbackorders"
    endif
firstrecord
brange=""
zitem=""
zbackrunning="No"    
˛/BackordersRun/∫ˇ˛SelectAllBackordersˇsynchronize 
select Backorder≠""˛/SelectAllBackordersˇ˛exportbackordersˇwaswindow= info("WindowName") 
form=info("FormName")
YesNo "Did you synchronize?"
if clipboard()="Yes"
        if form="seedspagecheck"    
openfile "backorders"
oldwindow=info("windowname")
window waswindow
firstrecord
        back=""
            loop
        linenu=1
        loop
        back=back+extract(extract(Backorder,¶,linenu),¬,1)+¬+?(extract(extract(Backorder,¶,linenu),¬,2) CONTAINS "-",extract(extract(Backorder,¶,linenu),¬,2)[1,"-"][1,-2],extract(extract(Backorder,¶,linenu),¬,2))+¬+
          ?(extract(extract(Backorder,¶,linenu),¬,2) CONTAINS "-",extract(extract(Backorder,¶,linenu),¬,2)[-1,-1],"")+¬
          +extract(extract(Backorder,¶,linenu),¬,3)+¬+¬+extract(extract(Backorder,¶,linenu),¬,4)[1,-1]+¶
        linenu=linenu+1
        Until  extract(extract(Backorder,¶,linenu),¬,2)=""   
            downrecord
            until  info("Stopped") 
window oldwindow
openfile "&@back"
save
window waswindow
        endif
endif
˛/exportbackordersˇ˛set_mail_flag_to_Bˇvmail="BG"˛/set_mail_flag_to_Bˇ˛.find_group/¢ˇmessage "No longer in use? Replaced by set_mail_flag_to_B?"
stop
zcancel="No"
form= info("FormName") 
If info("selected") < info("records")
SelectAll
EndIf
YesNo "Are you STILL processing mail?"
if clipboard()="Yes"
vmail="pB"
else
vmail="uB"
endif
;message vmail
getscrap "What's the order number?"
Numb=val(clipboard())
find OrderNo=Numb OR «5SpareNumber»=Numb
if info("Found")
stop
else
YesNo "Nothing found, try again?"
    if clipboard()="Yes"
    call justlocate/4
    endif
endif

˛/.find_group/¢ˇ˛RunThisBackorderˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
            if Zip = 0 and «9SpareText» ≠ ""
            message "This is a Canadian order. It should not have any backorders. Talk to Ryan or Bernice. This macro will stop"
            stop
            endif
    if str(OrderNo) contains "." OR Order=""
    NoYes "Do you want to run all the backorders for this group?"
    if clipboard()="Yes"
    Call "RunGroupBackorders"
    stop
    endif
    endif
YesNo "Do you want to run the backorder for this order?"
if clipboard()="Yes"    
ono=OrderNo
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup
find OrderNo=ono
waswindow=info("WindowName") 
    if Backorder=""
    message "There is no backorder here to run!"
    stop
    endif
    If Backorder NOTCONTAINS ¬+"  0"+¬
    message "This backorder was recently run and either wasn't processed on the computer or failed to update the server properly. This process will now stop."
    stop
    endif
    
openfile zcomyear+"SeedsComments linked"
synchronize
save
closefile
openfile zcomyear+"SeedsComments"
openfile "&&"+zcomyear+"SeedsComments linked"
save
window "Hide This Window"
window waswindow

    brange="No"
    zbackrunning="Yes"
Select OrderNo=ono
Call ".backordercomments"
Call "reworkchanges/®"
    openform "BackorderPicksheet"
    print ""
    CloseWindow
if arraysize(Backorder,¶)>35
message "This order has more than 35 lines of backorder. Please check with Ryan. They may not get all their backorders."
endif        
selectall
Find OrderNo=ono
brange=""
zbackrunning="No"
endif
    endif
˛/RunThisBackorderˇ˛RunGroupBackordersˇzitem=""
waswindow=info("WindowName") 
form= info("FormName")
Numb=int(OrderNo) ;6-2 change
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
Select int(OrderNo)=Numb AND Backorder≠"" AND FillDate≠today()  ;6-2 change
    if str(OrderNo) notcontains "."
    message "This is not a group order. Check it out!"
    stop
    endif
    if  info("Selected") =  info("Records") 
    message "There are no backorders here to run."
    stop
    endif
    
openfile zcomyear+"SeedsComments linked"
synchronize
save
closefile
openfile zcomyear+"SeedsComments"
openfile "&&"+zcomyear+"SeedsComments linked"
save
window "Hide This Window"
window waswindow

YesNo "Do you want to run ALL this group's backorders?"
    If clipboard()="Yes"
    brange="No"
selectwithin Backorder CONTAINS ¬+"  0"+¬     
zbackrunning="Yes"    
FirstRecord
Loop    
    If Backorder NOTCONTAINS ¬+"  0"+¬  ;this seems unnecessary-see above
    message "Order "+str(OrderNo)+"'s backorder was recently run and either wasn't processed or failed to update the server properly. It will not be run again. Check it out!"
    «10SpareText»="ran"
    endif
If Backorder CONTAINS ¬+"  0"+¬
Call ".backordercomments"
Call "reworkchanges/®"
endif
DownRecord
Until  info("Stopped") 
    SelectWithin «10SpareText»≠"ran"  ;this seems unnecessary-see above
    openform "BackorderPicksheet"
    print ""
    CloseWindow
if arraysize(Backorder,¶)>35
message "This order has more than 35 lines of backorder. Please check with Ryan. They may not get all their backorders."
endif        
window waswindow
firstrecord
zitem=""
    brange=""
    zbackrunning="No"    
    endif
˛/RunGroupBackordersˇ˛findGroupUndoneˇform= info("FormName") 
If info("selected") < info("records")
SelectAll
EndIf
vmail=""
getscrap "What's the order number?"
Numb=val(clipboard())
select val(str(OrderNo)[1,5))=Numb AND Backorder NOTCONTAINS ¬+"  0"+¬ AND Backorder≠"" 
if info("Found")
stop
else
Message "Nothing found, do you really know where you're going?"
endif˛/findGroupUndoneˇ˛BackorderFinderˇselectall
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
GetScrapOK "what item on B/O? (remember space for 3-digits)"
Select Backorder contains clipboard()˛/BackorderFinderˇ˛backorderlistˇ;this macro selects for the backorders that still have items on backorder
;the reverse are the backorders that are complete
;remember to use a space in front of 3-digit item numbers
;message "This list is not current. The macro will stop at this point."
;stop

Select Order = ""  ;selects to avoid empty orders and group cover sheets and starts the selection

       ; selects to avoid individual orders or group pieces
       NoYes "Do you want to avoid either individual orders or group pieces?"
       if clipboard()="Yes"
         NoYes "Do you want to avoid individual orders?"
         if clipboard()="Yes"
         selectAdditional str(OrderNo) notcontains "."
         else
         NoYes "Do you want to avoid group pieces?"
         if clipboard()="Yes"
         selectAdditional str(OrderNo) contains "."
         endif
         endif
       endif  
       if clipboard()="No"
       clipboard()=""
       endif

;SelectAdditional Backorder contains " 205-C"
;Select Backorder contains " 205" AND (Subs="N" OR Subs="O")
;Select Backorder contains "4682" AND (Subs="Y" OR Subs="C")
;SelectAdditional Backorder contains "2051" AND (Subs="N" OR Subs="O" OR Subs="C")
;SelectAdditional Backorder contains "4059-A" and (Subs notcontains "Y" OR Subs notcontains "C")
;SelectAdditional Backorder contains "1683" AND FirstFillDate>date("1/24/17")

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
˛/backorderlistˇ˛BackorderLineCountˇselectall
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
GetScrapOK "Minimum lines you want to select for?"
Select arraysize(Backorder,¶)≥val(clipboard())+1
if  info("Selected") =  info("Records") 
message "Nothing found with that many backorders."
endif˛/BackorderLineCountˇ˛more_than_one_lineˇmessage "This macro selects backorders with 2 or more lines within the current selection."
Selectwithin arraysize(Backorder,¶)≥3
˛/more_than_one_lineˇ˛find_southern_bo'sˇ    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
field OrderNo
sortup
select Zip ≥ 27000 AND Zip ≤ 42999  ;NC,SC,GA,FL,AL,TN,MS,KY
SelectAdditional Zip ≥ 63000 AND Zip ≤ 65999  ;MO
SelectAdditional Zip ≥ 70000 AND Zip ≤ 79999  ;LA,AR,OK,TX
SelectAdditional Zip ≥ 84000 AND Zip ≤ 89999  ;UT,AZ,NM,NV
SelectAdditional Zip ≥ 90000 AND Zip ≤ 98999  ;CA,OR,HI,WA
selectwithin Backorder ≠ ""˛/find_southern_bo'sˇ˛RunBackorderBatchˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
waswindow=info("WindowName") 
YesNo "Did you synchronize?"
    if clipboard()="Yes"
openfile zcomyear+"SeedsComments linked"
synchronize
save
closefile
openfile zcomyear+"SeedsComments"
openfile "&&"+zcomyear+"SeedsComments linked"
save
window "Hide This Window"
window waswindow
field OrderNo
sortup
firstrecord
YesNo "Do you want to run the backorders for all the selected orders?"
    if clipboard()="Yes"
    brange="No"
selectwithin Backorder CONTAINS ¬+"  0"+¬   
zbackrunning="Yes"      
loop
if Backorder=""
message "Order "+str(OrderNo)+" has no backorder to run!"
endif
    If Backorder NOTCONTAINS ¬+"  0"+¬  ;this seems unnecessary-see above
    message "Order "+str(OrderNo)+"'s backorder was recently run and either wasn't processed or failed to update the server properly. It will not be run again. Check it out!"
    «10SpareText»="ran"
    endif
If Backorder CONTAINS ¬+"  0"+¬
Call ".backordercomments"
Call "reworkchanges/®"
endif
downrecord
until  info("Stopped")
    SelectWithin «10SpareText»≠"ran"   ;this seems unnecessary-see above
    openform "BackorderPicksheet"
    print ""
    CloseWindow
if arraysize(Backorder,¶)>35
message "This order has more than 35 lines of backorder. Please check with Ryan. They may not get all their backorders."
endif        
    YesNo "Do you want to export this set to generate a commodity list?"
    if clipboard()="Yes"
    call "exportbackorders"
    endif
firstrecord
    brange=""
    zbackrunning="No"    
    endif
    endif
    endif
˛/RunBackorderBatchˇ˛ViewBackorderˇ    if  info("Windows") notcontains "ViewBackorder"
    WindowBox "30 640 600 1120 noPalette noHorzScroll"
    openform "ViewBackorder"
    else
    goform "ViewBackorder"
    zoomwindow 30,640,550,490,"noPalette noHorzScroll"
    endif
˛/ViewBackorderˇ˛.backorder_recordˇif Backorder_History=""
    Backorder_History=datepattern(today(),"mm/dd/yy")+¶+Backorder
else
    Backorder_History=datepattern(today(),"mm/dd/yy")+¶+Backorder+rep("_",20)+¶+¶+Backorder_History
endif

;archived old dating
;if Backorder_History=""
;    Backorder_History=?(datepattern(today(),"Day") contains "Fri",
;    datepattern(today()+3,"mm/dd/yy"),
;    datepattern(today()+1,"mm/dd/yy"))+¶+Backorder
;else
;    Backorder_History=?(datepattern(today(),"Day") contains "Fri",
;    datepattern(today()+3,"mm/dd/yy"),
;    datepattern(today()+1,"mm/dd/yy"))+¶+Backorder+rep("_",20)+¶+¶+Backorder_History
;endif˛/.backorder_recordˇ˛.backorder_countˇ;this checks if the backorder in progress has the same number of items on backorder as lines to keep it off the backorder history and manifest
zlinect=0
zbackrun="Yes"
zleftbo=""
zleftct=0
linenu=1

zlinect = arraysize(Backorder,¶)-1
;message zlinect

loop
linenu=arraysearch(Backorder,"*backorder*", linenu,¶)
stoploopif linenu=0
zleftitem=extract(Backorder,¶,linenu)+¶
zleftbo=zleftbo+zleftitem
linenu=linenu+1
until info("empty")
zleftct = arraysize(zleftbo,¶)-1
;message zleftct

if zlinect-zleftct=0
zbackrun="No"
else
zbackrun="Yes"
endif

;message zbackct

;if zlinect=zleftct
;message "yo"
;else
;message "no yo"
;endif
˛/.backorder_countˇ˛viewBackorderHistoryˇ        if  info("Windows") notcontains "ViewBackHistory"
        WindowBox "580 640 850 1120 noPalette noHorzScroll"
        openform "ViewBackHistory"
        else
        goform "ViewBackHistory"
        zoomwindow 580,640,280,490,"noPalette noHorzScroll"
        endif
˛/viewBackorderHistoryˇ˛fix backorder historyˇWindowBox "300 340 850 820 noPalette noHorzScroll"
openform "FixBackHistory"
EditCellStop
˛/fix backorder historyˇ˛select_unprintedˇselectwithin Backorder CONTAINS ¬+"  0"+¬   
˛/select_unprintedˇ˛find_completesˇselectall
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
field OrderNo
sortup
message "This macro finds backorders that can be completed using the backorder list of what's not available as a starting point, so make sure it's up-to-date."
NoYes "Do you still want to look for backorders that can be completed?"
If clipboard()="Yes"
call "backorderlist"
selectreverse
selectwithin Backorder ≠ ""
selectwithin Backorder CONTAINS ¬+"  0"+¬ 
endif˛/find_completesˇ˛FindUndoneˇSynchronize
field OrderNo
sortup
ztime=now() ;resets sync timer
select Backorder NOTCONTAINS ¬+"  0"+¬ AND Backorder≠""
if  info("Selected") = info("Records") 
message "Nothing found."
endif˛/FindUndoneˇ˛.backorderˇbackline=""
added=""
linenu=1
loop
linenu=arraysearch(Order,"*backorder*", linenu,¶)
stoploopif linenu=0
backline=extract(Order,¶,linenu)+¶
added=added+backline
linenu=linenu+1
until info("empty")
Backorder=added
if added=""
message "There is no backorder"
stop
endif˛/.backorderˇ˛.backordercommentsˇzitem=""
zsize=""
zcomment=""
zname=""
subnum=""
linenu=1
zprice=0
zfill=0
ztot=0
sub=Subs
loop
zitem=extract(extract(Backorder,¶,linenu),¬,2)[1,"-"][1,-2]
stoploopif zitem=""
zsize=extract(extract(Backorder,¶,linenu),¬,2)[-1,-1]

zcomment=?(zsize CONTAINS "A",lookup(zcomyear+"SeedsComments","Item",zitem,"comA","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjA","",0),
                 ?(zsize CONTAINS "B",lookup(zcomyear+"SeedsComments","Item",zitem,"comB","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjB","",0),
                 ?(zsize CONTAINS "C",lookup(zcomyear+"SeedsComments","Item",zitem,"comC","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjC","",0),
                 ?(zsize CONTAINS "D",lookup(zcomyear+"SeedsComments","Item",zitem,"comD","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjD","",0),
                 ?(zsize CONTAINS "E",lookup(zcomyear+"SeedsComments","Item",zitem,"comE","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjE","",0),
                 ?(zsize CONTAINS "K",lookup(zcomyear+"SeedsComments","Item",zitem,"comK","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjK","",0),
                 ?(zsize CONTAINS "L",lookup(zcomyear+"SeedsComments","Item",zitem,"comL","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjL","",0),
                 ?(zsize CONTAINS "X",lookup(zcomyear+"SeedsComments","Item",zitem,"comX","",0)
                        +" "+lookup(zcomyear+"SeedsComments","Item",zitem,"adjX","",0),""))))))))
if lookup(zcomyear+"SeedsComments","Item",zitem,"orders_only","",0)="x" and zcomment contains "sub" ;resetting subs for new orders only
zcomment="backorder"
endif
zcomment=strip(zcomment)

vcode=lookup(zcomyear+"SeedsComments","Item",zitem,"special","",0)
            if zcomment CONTAINS "sub"
                if sub="N"
                zcomment="sold out"
                endif
                    if sub="C"
                    if vcode≠"CO"
                    zcomment="sold out"
                    endif  
                    endif
                if sub="O"
                if vcode≠"OO" AND vcode≠"" ;this covers items that are not organic
                zcomment="sold out"
                endif 
                endif    
            endif

call .commentAdjustments

if zcomment="here"
zcomment="backorder"
endif
             
zname=lookup(zcomyear+"SeedsComments","Item",zitem,"Description","",0)
if zcomment CONTAINS "sub"
subnum=?(zcomment notcontains "-",striptonum(zcomment),
                ?(zcomment contains "-",striptonum(zcomment["-",-1]),""))
zname="sub-"+lookup(zcomyear+"SeedsComments","Item",subnum,"Description","",0)
endif

zprice=val(extract(extract(Backorder,¶,linenu),¬,6)[1,-1])
if zcomment CONTAINS "sub"
    if lookup(zcomyear+"SeedsComments","Item",subnum,"price"+zsize,"",0) < zprice
    zprice=lookup(zcomyear+"SeedsComments","Item",subnum,"price"+zsize,"",0)
    endif
endif
if zcomment CONTAINS "limit" ;;; OR zcomment CONTAINS "sub" - now using lookup above for subs
zprice=?(zsize CONTAINS "A",lookup(zcomyear+"SeedsComments","Item",zitem,"priceA","",0),
          ?(zsize CONTAINS "B",lookup(zcomyear+"SeedsComments","Item",zitem,"priceB","",0),
          ?(zsize CONTAINS "C",lookup(zcomyear+"SeedsComments","Item",zitem,"priceC","",0),
          ?(zsize CONTAINS "D",lookup(zcomyear+"SeedsComments","Item",zitem,"priceD","",0),
          ?(zsize CONTAINS "E",lookup(zcomyear+"SeedsComments","Item",zitem,"priceE","",0),
          ?(zsize CONTAINS "K",lookup(zcomyear+"SeedsComments","Item",zitem,"priceK","",0),
          ?(zsize CONTAINS "L",lookup(zcomyear+"SeedsComments","Item",zitem,"priceL","",0),
          ?(zsize CONTAINS "X",lookup(zcomyear+"SeedsComments","Item",zitem,"priceX","",0),""))))))))
endif

;price adjustments go here

zfill=val(extract(extract(Backorder,¶,linenu),¬,3)["0-9",-1])
    if zcomment CONTAINS "out"
    zfill=0
    endif
ztot=pattern(zfill*zprice,"#.##")
    if zcomment=" " and vcode="O"
    zcomment="organic"
    endif
Backorder=arraychange(Backorder,extract(extract(Backorder,¶,linenu),¬,1)
                    +¬+extract(extract(Backorder,¶,linenu),¬,2)
                    +¬+extract(extract(Backorder,¶,linenu),¬,3)
                    +?(zcomment contains "out" OR zcomment contains "backorder" 
                        OR zcomment contains "member" OR zcomment contains "sub","",rep(chr(95),3))
                    +¬+?(zcomment contains "out" OR zcomment contains "backorder" 
                        OR zcomment contains "member" OR zcomment contains "sub",zcomment+rep(chr(32),12-length(zcomment)),
                       zcomment+rep(chr(32),9))
                    +¬+rep(chr(32),6-length(str(pattern(zprice,"#.##"))))+str(pattern(zprice,"#.##"))
                    +¬+rep(chr(32),7-length(str(ztot)))+str(ztot)
                    +¬+zname,linenu,¶)

linenu=linenu+1
until info("empty")
;stop
call ".orderchanger"
˛/.backordercommentsˇ˛.commentAdjustmentsˇ;don't forget the leading space on 3-digit item numbers
;remember: no dash before size letters
;message "The adjustment list is not up to date. The macro will stop."
;stop

;if zitem=" 205"
;zcomment="backorder"
;endif

;if zitem="2265"
;zcomment=""
;endif

;if zitem="1615"
;if zsize="A"
;zcomment=""
;endif
;if zsize="B"
;zcomment=""
;endif

;if zsize="C"
;zcomment=""
;endif

;if zsize="D"
;zcomment=""
;endif
;endif

;if zitem="3645"
;if sub="N" OR sub="C"
;zcomment="sold out"
;endif
;endif

;this is a kill switch for anything on backorder
;if zcomment contains "backorder"
;zcomment="sold out
;endif

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;


˛/.commentAdjustmentsˇ˛.orderchangerˇlinenu=1
zitem=""
zcomment=""
zline=1
loop
zitem=extract(extract(Backorder,¶,linenu),¬,2)
stoploopif zitem=""
zcomment=extract(extract(Backorder,¶,linenu),¬,4)[1,12]
zline=arraysearch(Order,"*"+zitem+"*",zline,¶)
stoploopif zline=0
Order=arraychange(Order,extract(extract(Backorder,¶,linenu),¬,1)
            +¬+extract(extract(Backorder,¶,linenu),¬,2)
            +¬+extract(extract(Backorder,¶,linenu),¬,3)
            +¬+extract(extract(Backorder,¶,linenu),¬,3)
            +¬+extract(extract(Backorder,¶,linenu),¬,4)
            +¬+extract(extract(Backorder,¶,linenu),¬,5)
            +¬+extract(extract(Backorder,¶,linenu),¬,6)
            +¬+extract(extract(Backorder,¶,linenu),¬,7),zline,¶)
linenu=linenu+1
zline=zline+1
until info("Empty")
˛/.orderchangerˇ˛--ˇ˛/--ˇ˛(doodah)ˇ˛/(doodah)ˇ˛ImportOrders/7ˇglobal checkprocedure, discprocedure, zmessage
    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
    if info("unixusername") CONTAINS "tyee" or info("unixusername") CONTAINS "gene" or info("unixusername") CONTAINS "Sensation"
    
checkprocedure=""
getproceduretext "",".membership_check",checkprocedure
discprocedure=""
getproceduretext "",".discount_check",discprocedure
zmessage=""
    
;field OrderNo
;sortup
    ;if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    ;endif
lastrecord
;Numb=OrderNo
Numb=val(info("serverrecordid"))
;message info("serverrecordid")
;stop

OpenFileDialog folder,file,type,""
YesNo "Is - "+file+" - the file you want to append?"

If clipboard()="Yes"
openfile folderpath(folder)+file
gosheet
ZoomWindow 25,692,448,621,""
copycell
pastecell
makenewprocedure ".membership_check",""
openprocedure ".membership_check"
setproceduretext checkprocedure
closewindow
save
makenewprocedure ".discount_check",""
openprocedure ".discount_check"
setproceduretext discprocedure
closewindow
save

;now check for members
call .membership_check

;check for previous discounts
call .discount_check

;append to composite order file
openfile "seeds all"
openfile "++"+folderpath(folder)+file
save
closefile
openfile folderpath(folder)+file
closefile

;append to tally, flag office order file and select orders to check
window waswindow
openfile "++"+folderpath(folder)+file
;openfile "++"+folder+":"+file

executeapplescriptformula {tell application "Finder"
set var1 to "«file»"
	set name of document file var1 of folder "zoffice files" of folder "44tally.seeds" of folder "Desktop" of folder "tyee" of folder "Users" of startup disk to var1 & "√"
end tell}
;closefile

select val(info("serverrecordid"))>Numb AND Order CONTAINS "201)"
; OR (val(info("serverrecordid"))>Numb AND City CONTAINS "miami")
    call select_unprinted_racks
        if  info("Selected") <  info("Records") 
    
        tallmessage "The selected order(s) in the tally have a seed rack or more than 200 items or are from Miami."+¶+¶+zmessage
        else
        tallmessage zmessage
        endif

endif

    endif
    endif
    
˛/ImportOrders/7ˇ˛PrintThisPicksheetˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
zlarge=0
ono=OrderNo
            if Zip = 0 and «9SpareText» ≠ ""
            message "This is a Canadian order. You cannot work on it here. Use the Canadian setup. This macro will stop"
            stop
            endif
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup
Select OrderNo=ono
oldfile=  info("DatabaseName") 
waswindow=info("windowname")
form=info("FormName")
rayg=""
if  info("Selected") =1
    if ShipCode="D"
    message "This order is being held in the office for lack of payment. It cannot be printed."
    stop
    endif
If Order contains "0000" OR Order = "" OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderComments CONTAINS "empty order")
    OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderNo>1000000) OR PickSheet≠"" ;99000) OR PickSheet≠"" 9-3 fix?
    message "There's nothing here to print or the picksheet has already been printed for this order or this is a group cover sheet."
    selectall
    Find OrderNo=ono
    stop
endif
        if form="seedspagecheck"    
        openfile zcomyear+"SeedsComments linked"
        ReSynchronize
        save
        closefile
        openfile zcomyear+"SeedsComments"
        openfile "&&"+zcomyear+"SeedsComments linked"
        endif
save
window "Hide This Window"
window waswindow
call ".comments"

if  PickSheet≠""
firstrecord
        PrintUsingForm "", "seedspicksheet"
print ""
    Selectwithin PickSheet CONTAINS "41)"
        If  info("Empty")
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet2"
        print ""
   endif
    Selectwithin PickSheet CONTAINS "96)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet3"
        print ""
   endif
    Selectwithin PickSheet CONTAINS "151)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet4"
        print ""
    endif
    skip:
    
    selectall
    Find OrderNo=ono
        if OrderComments≠"" OR PickSheet CONTAINS "201)" OR Notes1≠""
        if PickSheet CONTAINS "201)"
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!
Also, one or more order(s) has more than 200 items. Check with Bernice or Carol about this!"
       endif
        if PickSheet NOTCONTAINS "201)"
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!"
       endif
       endif
else 
message "Nothing to print!"
endif  
endif     
endif
  ˛/PrintThisPicksheetˇ˛no_stats/6ˇzstat=""
    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
NoYes "Are you sure you want to set the stats for this order?"
if clipboard()="Yes"
zstat="Yes"
Checker="Nobody"
Puller="Everybody"
Notes1=Notes1+" "+"This order had no stats but was shipped. Set stats "+datepattern(today(),"mm/dd/yy")+"."
call .done
endif
    endif
    zstat=""
˛/no_stats/6ˇ˛.membership_checkˇ;this macro gets pasted into each incoming office file
;global zmemEmail, zmemPhone, zmemCon, zmemAdd, zmemNo, zmemGroup, zmembers, zmemberwindow, zadjlist
zmemEmail=""
zmemPhone=""
zmemCon=""
zmemAdd=""
zmemNo=""
zmemGroup=""
zmembers=""
zmemberwindow=""
zadjlist=""
field OrderNo
sortup
;zmemberwindow="Gene2 miniHD:Users:gene2:Desktop:"+zyear+"tally.seeds:member_list"
zmemberwindow="tyee-miniHD:Users:tyee:Desktop:"+zyear+"tally.seeds:member_list"
;zmemberwindow="sensation:Users:statice:Desktop:"+zyear+"tally.seeds:member_list"
    if info("files") notcontains "member_list"
    openfile zmemberwindow
    openfile folderpath(folder)+file
    endif
;stop    
;zmembers=lookup("member_list",Con,Con,«Mem?»,"N",0)
arrayselectedbuild zmemNo,¶,"",?(MemDisc=0 AND ShipCode≠"D"AND
                                lookup("member_list",«C#»,«C#»,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",«C#»,«C#»,«Mem?»,"N",0),"")

arrayselectedbuild zmemCon,¶,"",?(MemDisc=0 AND ShipCode≠"D"AND
                                lookup("member_list",Con,Con,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",Con,Con,«Mem?»,"N",0),"")
                                
arrayselectedbuild zmemGroup,¶,"",?(MemDisc=0 AND ShipCode≠"D"AND
                                lookup("member_list",Group,Group,«Mem?»,"N",0)="Y" AND Group≠"" AND lookup("member_list",Zip,Zip,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",Group,Group,«Mem?»,"N",0),"")
                                
arrayselectedbuild zmemAdd,¶,"",?(MemDisc=0 AND ShipCode≠"D"AND
                                lookup("member_list",MAd,MAd,«Mem?»,"N",0)="Y" AND lookup("member_list",Zip,Zip,«Mem?»,"N",0)="Y"
                                 AND lookup("member_list",SAd,SAd,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",MAd,MAd,«Mem?»,"N",0),"")
                                
arrayselectedbuild zmemEmail,¶,"",?(MemDisc=0 AND ShipCode≠"D"AND
                                lookup("member_list",email,Email,«Mem?»,"N",0)="Y" AND Email≠""
                                ,str(OrderNo)+"-"+lookup("member_list",email,Email,«Mem?»,"N",0),"")

arrayselectedbuild zmemPhone,¶,"",?(MemDisc=0 AND ShipCode≠"D"AND
                                lookup("member_list",phone,?(length(Telephone)>12 AND Telephone CONTAINS "-",Telephone[1,12],
                      ?(length(Telephone)=12 AND Telephone NOTCONTAINS "-", Telephone[1,3]+"-"+Telephone[5,7]+"-"+Telephone[9,12],
                      ?(length(Telephone)=10, Telephone[1,3]+"-"+Telephone[4,6]+"-"+Telephone[7,10],
                      Telephone))),«Mem?»,"N",0)="Y" AND Telephone≠""
                                ,str(OrderNo)+"-"+lookup("member_list",phone,?(length(Telephone)>12 AND Telephone CONTAINS "-",Telephone[1,12],
                      ?(length(Telephone)=12 AND Telephone NOTCONTAINS "-", Telephone[1,3]+"-"+Telephone[5,7]+"-"+Telephone[9,12],
                      ?(length(Telephone)=10, Telephone[1,3]+"-"+Telephone[4,6]+"-"+Telephone[7,10],
                      Telephone))),«Mem?»,"N",0),"")

zmembers= "MEMBER POSSIBILITIES:"+¶+"cust num"+¶+zmemNo+¶+"name"+¶+zmemCon+¶+"group"+¶+zmemGroup+¶+"address"
                    +¶+zmemAdd+¶+"email"+¶+zmemEmail+¶+"phone"+¶+zmemPhone
                                                           
;arraybuild zmembers,¶,"",?(MemDisc=0,OrderNo,"")
zadjlist=zmembers
clipboard()=zadjlist
;    if zmemNo="" AND zmemCon="" AND zmemGroup="" AND zmemAdd="" AND zmemEmail="" AND zmemPhone=""
;    message "nothing found, YAY!"
;    else
;    tallmessage zadjlist
;    endif
;lookupall("member_list",Con,Con,«Mem?»,¶)
    openfile zmemberwindow
    closefile˛/.membership_checkˇ˛.discount_checkˇ;this macro gets pasted into each incoming office file
;global zdiscEmail, zdiscPhone, zdiscCon, zdiscAdd, zdiscNo, zdiscGroup, zdiscounts, zdiscountwindow, zadjlist
zdiscEmail=""
zdiscPhone=""
zdiscCon=""
zdiscAdd=""
zdiscNo=""
zdiscGroup=""
zdiscounts=""
zdiscountwindow=""
field OrderNo
sortup
;zdiscountwindow="Gene2 miniHD:Users:gene2:Desktop:"+zyear+"tally.seeds:seeds all"
zdiscountwindow="tyee-miniHD:Users:tyee:Desktop:"+zyear+"tally.seeds:seeds all"
;zdiscountwindow="sensation:Users:statice:Desktop:"+zyear+"tally.seeds:seeds all"
    if info("files") notcontains "seeds all"
    openfile zdiscountwindow
    selectall
    openfile folderpath(folder)+file
    endif
;stop    

;zdiscounts=lookup("seeds all",Con,Con,Discount,0,0)
arrayselectedbuild zdiscNo,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND
                                lookup("seeds all",«C#»,«C#»,Discount,0,0)>Discount AND «C#»≠0
                                ,str(OrderNo)+"-"+str(lookup("seeds all",«C#»,«C#»,Discount,0,0)),"")

arrayselectedbuild zdiscCon,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND
                                lookup("seeds all",Con,Con,Discount,0,0)>Discount
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Con,Con,Discount,0,0)),"")
                                
arrayselectedbuild zdiscGroup,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND
                                lookup("seeds all",Group,Group,Discount,0,0)>Discount AND Group≠"" AND lookup("seeds all",Zip,Zip,Discount,0,0)>Discount
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Group,Group,Discount,0,0)),"")
                                
arrayselectedbuild zdiscAdd,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND
                                lookup("seeds all",MAd,MAd,Discount,0,0)>Discount AND lookup("seeds all",Zip,Zip,Discount,0,0)>Discount
                                 AND lookup("seeds all",SAd,SAd,Discount,0,0)>Discount
                                ,str(OrderNo)+"-"+str(lookup("seeds all",MAd,MAd,Discount,0,0)),"")
                                
arrayselectedbuild zdiscEmail,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND
                                lookup("seeds all",Email,Email,Discount,0,0)>Discount AND Email≠""
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Email,Email,Discount,0,0)),"")

arrayselectedbuild zdiscPhone,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND
                                lookup("seeds all",Telephone,Telephone,Discount,0,0)>Discount AND Telephone≠""
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Telephone,Telephone,Discount,0,0)),"")

zdiscounts= "PREVIOUS DISCOUNT POSSIBILITIES:"+¶+"cust num"+¶+zdiscNo+¶+"name"+¶+zdiscCon+¶+"group"+¶+zdiscGroup+¶+"address"
                    +¶+zdiscAdd+¶+"email"+¶+zdiscEmail+¶+"phone"+¶+zdiscPhone
                                                           
zadjlist=zadjlist+¶+zdiscounts
clipboard()=zadjlist
    if zmemNo="" AND zmemCon="" AND zmemGroup="" AND zmemAdd="" AND zmemEmail="" AND zmemPhone="" AND
         zdiscNo="" AND zdiscCon="" AND zdiscGroup="" AND zdiscAdd="" AND zdiscEmail="" AND zdiscPhone=""
    zmessage = "no member discounts or previous discounts were found, YAY!"
    else
    zmessage = zadjlist
    endif
˛/.discount_checkˇ˛member_check_tallyˇ;global zmemEmail, zmemPhone, zmemCon, zmemAdd, zmemNo, zmemGroup, zmembers
zmemEmail=""
zmemPhone=""
zmemCon=""
zmemAdd=""
zmemNo=""
zmemGroup=""
zmembers=""
selectall
synchronize
field OrderNo
sortup
waswindow=zyear+"seedstally:seedspagecheck"
    if info("files") notcontains "member_list"
    openfile "member_list"
    window waswindow
    endif
    
;zmembers=lookup("member_list",Con,Con,«Mem?»,"N",0)
arrayselectedbuild zmemNo,¶,"",?(MemDisc=0 AND ShipCode≠"D" AND «not_member»≠"x" AND
                                lookup("member_list",«C#»,«C#»,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",«C#»,«C#»,«Mem?»,"N",0),"")

arrayselectedbuild zmemCon,¶,"",?(MemDisc=0 AND ShipCode≠"D" AND «not_member»≠"x" AND
                                lookup("member_list",Con,Con,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",Con,Con,«Mem?»,"N",0),"")
                                
arrayselectedbuild zmemGroup,¶,"",?(MemDisc=0 AND ShipCode≠"D" AND «not_member»≠"x" AND
                                lookup("member_list",Group,Group,«Mem?»,"N",0)="Y" AND Group≠"" AND lookup("member_list",Zip,Zip,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",Group,Group,«Mem?»,"N",0),"")
                                
arrayselectedbuild zmemAdd,¶,"",?(MemDisc=0 AND ShipCode≠"D" AND «not_member»≠"x" AND
                                lookup("member_list",MAd,MAd,«Mem?»,"N",0)="Y" AND lookup("member_list",Zip,Zip,«Mem?»,"N",0)="Y"
                                 AND lookup("member_list",SAd,SAd,«Mem?»,"N",0)="Y"
                                ,str(OrderNo)+"-"+lookup("member_list",MAd,MAd,«Mem?»,"N",0),"")
                                
arrayselectedbuild zmemEmail,¶,"",?(MemDisc=0 AND ShipCode≠"D" AND «not_member»≠"x" AND
                                lookup("member_list",email,Email,«Mem?»,"N",0)="Y" AND Email≠""
                                ,str(OrderNo)+"-"+lookup("member_list",email,Email,«Mem?»,"N",0),"")

arrayselectedbuild zmemPhone,¶,"",?(MemDisc=0 AND ShipCode≠"D" AND «not_member»≠"x" AND
                                lookup("member_list",phone,?(length(Telephone)>12 AND Telephone CONTAINS "-",Telephone[1,12],
                      ?(length(Telephone)=12 AND Telephone NOTCONTAINS "-", Telephone[1,3]+"-"+Telephone[5,7]+"-"+Telephone[9,12],
                      ?(length(Telephone)=10, Telephone[1,3]+"-"+Telephone[4,6]+"-"+Telephone[7,10],
                      Telephone))),«Mem?»,"N",0)="Y" AND Telephone≠""
                                ,str(OrderNo)+"-"+lookup("member_list",phone,?(length(Telephone)>12 AND Telephone CONTAINS "-",Telephone[1,12],
                      ?(length(Telephone)=12 AND Telephone NOTCONTAINS "-", Telephone[1,3]+"-"+Telephone[5,7]+"-"+Telephone[9,12],
                      ?(length(Telephone)=10, Telephone[1,3]+"-"+Telephone[4,6]+"-"+Telephone[7,10],
                      Telephone))),«Mem?»,"N",0),"")

zmembers= "cust num"+¶+zmemNo+¶+"name"+¶+zmemCon+¶+"group"+¶+zmemGroup+¶+"address"
                    +¶+zmemAdd+¶+"email"+¶+zmemEmail+¶+"phone"+¶+zmemPhone
                                                           
;arraybuild zmembers,¶,"",?(MemDisc=0,OrderNo,"")
clipboard()=zmembers
    if zmemNo="" AND zmemCon="" AND zmemGroup="" AND zmemAdd="" AND zmemEmail="" AND zmemPhone=""
    message "nothing found, YAY!"
    else
    tallmessage zmembers
    endif
;lookupall("member_list",Con,Con,«Mem?»,¶)
˛/member_check_tallyˇ˛change_membershipˇYesNo "Do You Want To Change The Membership Status Of This Order?"
if clipboard() = "Yes"
    if MemDisc > 0
    YesNo "This person is already a member! Should they NOT be a member?"
        if clipboard() = "Yes"
        MemDisc = 0
        endif
        message "Member Status has been changed, but the Totals have not been re-worked."
        stop
    endif
    if MemDisc = 0 and «3SpareText»≠"x"
        YesNo "Should this person be a member?"
            if clipboard() = "Yes"
            MemDisc = .01
            endif
     endif
    if MemDisc = 0 and «3SpareText»="x"
        message "This has been flagged as not a member, usually they've used 
        someone else's catalog. Please check. Nothing has been changed."
        stop
     endif
message "Member Status has been changed, but the Totals have not been re-worked."
endif
˛/change_membershipˇ˛discount_check_tallyˇ;global zdiscEmail, zdiscPhone, zdiscCon, zdiscAdd, zdiscNo, zdiscGroup, zdiscounts, zdiscountwindow, zadjlist
zdiscEmail=""
zdiscPhone=""
zdiscCon=""
zdiscAdd=""
zdiscNo=""
zdiscGroup=""
zdiscounts=""
zdiscountwindow=""
field OrderNo
sortup
waswindow=zyear+"seedstally:seedspagecheck"
    if info("files") notcontains "seeds all"
    openfile "seeds all"
    window waswindow
    endif


;zdiscounts=lookup("seeds all",Con,Con,Discount,0,0)
arrayselectedbuild zdiscNo,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND «no_discount»≠"x" AND
                                lookup("seeds all",«C#»,«C#»,Discount,0,0)>Discount
                                ,str(OrderNo)+"-"+str(lookup("seeds all",«C#»,«C#»,Discount,0,0)),"")

arrayselectedbuild zdiscCon,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND «no_discount»≠"x" AND
                                lookup("seeds all",Con,Con,Discount,0,0)>Discount
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Con,Con,Discount,0,0)),"")
                                
arrayselectedbuild zdiscGroup,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND «no_discount»≠"x" AND
                                lookup("seeds all",Group,Group,Discount,0,0)>Discount AND Group≠"" AND lookup("seeds all",Zip,Zip,Discount,0,0)>Discount
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Group,Group,Discount,0,0)),"")
                                
arrayselectedbuild zdiscAdd,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND «no_discount»≠"x" AND
                                lookup("seeds all",MAd,MAd,Discount,0,0)>Discount AND lookup("seeds all",Zip,Zip,Discount,0,0)>Discount
                                 AND lookup("seeds all",SAd,SAd,Discount,0,0)>Discount
                                ,str(OrderNo)+"-"+str(lookup("seeds all",MAd,MAd,Discount,0,0)),"")
                                
arrayselectedbuild zdiscEmail,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND «no_discount»≠"x" AND
                                lookup("seeds all",Email,Email,Discount,0,0)>Discount AND Email≠""
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Email,Email,Discount,0,0)),"")

arrayselectedbuild zdiscPhone,¶,"",?(ShipCode≠"D"AND Subtotal≥50 AND «no_discount»≠"x" AND
                                lookup("seeds all",Telephone,Telephone,Discount,0,0)>Discount AND Telephone≠""
                                ,str(OrderNo)+"-"+str(lookup("seeds all",Telephone,Telephone,Discount,0,0)),"")

zdiscounts= "cust num"+¶+zdiscNo+¶+"name"+¶+zdiscCon+¶+"group"+¶+zdiscGroup+¶+"address"
                    +¶+zdiscAdd+¶+"email"+¶+zdiscEmail+¶+"phone"+¶+zdiscPhone
                                                           
clipboard()=zdiscounts
    if zdiscNo="" AND zdiscCon="" AND zdiscGroup="" AND zdiscAdd="" AND zdiscEmail="" AND zdiscPhone=""
    message "no previous discounts were found, YAY!"
    else
    tallmessage zdiscounts
    endif
˛/discount_check_tallyˇ˛change_discountˇ            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
getscrap "what's the new discount? (decimal, please)"
Discount=val(clipboard())
    if PickSheet≠""
call ".retotal"
    endif
zcanada="No"˛/change_discountˇ˛change_subˇGetScrapOK "Change Subs To (Y, N, C or O)?"
If upper(clipboard()) = "Y" or upper(clipboard()) = "N" or upper(clipboard()) = "C" or upper(clipboard()) = "O" or clipboard()=""
Subs=upper(clipboard())
else
message "This is not a sub code."
endif˛/change_subˇ˛sync & sortup/0ˇono=OrderNo
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup
find OrderNo=ono˛/sync & sortup/0ˇ˛orders_pulledˇlocal zorderdate, zbackorderdate
selectall
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif

GetScrap "What date are you interested in? (e.g. 1/30/16)"
zorderdate=datepattern(date(clipboard()),"mm/dd/yy")

;I think this is out of date 1-25-16
;if datepattern(date(clipboard()),"Day") contains "Fri"
;zbackorderdate=datepattern(date(clipboard())+3,"mm/dd/yy")
;message zbackorderdate
;else
;zbackorderdate=datepattern(date(clipboard()),"mm/dd/yy")
;endif

;message zorderdate+¶+zbackorderdate
;Select FirstFillDate = date(zorderdate) and Backorder_History notcontains zbackorderdate and str(OrderNo)  NOTCONTAINS "."
Select FirstFillDate = date(zorderdate) and str(OrderNo)  NOTCONTAINS "."
˛/orders_pulledˇ˛select_printed_racksˇselectall
select (Order contains "5951" OR Order contains "5952" OR Order contains "5953" 
    OR Order contains "5954" OR Order contains "5955" OR Order contains "5956") and PickSheet≠""
 if  info("Selected") <  info("Records") 
stop
 else
 message "Nothing selected"
 endif
˛/select_printed_racksˇ˛select_unprinted_racksˇselectall
select (Order contains "5951" OR Order contains "5952" OR Order contains "5953" 
    OR Order contains "5954" OR Order contains "5955" OR Order contains "5956") and PickSheet=""
 if  info("Selected") <  info("Records") 
stop
 else
 message "Nothing selected"
 endif
˛/select_unprinted_racksˇ˛unprinted_group_piecesˇ;this selects unprinted group pieces
selectall
synchronize
field OrderNo
sortup

YesNo "Are you searching for unprinted PAPER group pieces?"
if clipboard()="Yes"
selectwithin PickSheet="" AND str(OrderNo) contains "." AND OrderNo<20000
    if info("records")=info("selected") 
    message "nothing found"
    selectall
    endif
else
YesNo "Are you searching for unprinted INTERNET group pieces?"
    if clipboard()="Yes"
selectwithin PickSheet="" AND str(OrderNo) contains "." AND OrderNo>99000
        if info("records")=info("selected") 
        message "nothing found"
        selectall
        endif
    endif
endif˛/unprinted_group_piecesˇ˛find_new_soil_kitsˇselectall
select (Order contains "5965" and Status="") OR (Order contains "8194" and Status="")
if info("selected")=info("records")
message "no new soil test kits"
endif˛/find_new_soil_kitsˇ˛runpaperwork/çˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
if Order≠"" or vmail="BG"
call .done
else
message "Please click Backorder or Complete to update this order. This method doesn't work on group cover sheets."
stop
endif
    endif˛/runpaperwork/çˇ˛zeroAddPay1ˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
YesNo "Do you want to zero out AddPay1?"
if clipboard()="Yes"    
Paid=«1stPayment»
AddPay1=0
Message "Please check that the payment fields make sense. Also, does the mail manifest need adjusting? Thanks! Ryan"
endif
if «1stPayment»+AddPay1+AddPay2+AddPay3+AddPay4+AddPay5+AddPay6≠Paid
message "Amount Paid doesn't match the payment fields. Talk to Bernice if you're unsure what to do, it can avoid bookkeeping headaches. Thanks. Ryan"
endif
    endif˛/zeroAddPay1ˇ˛OfficeInfo/øˇNumb=OrderNo
OpenFile "seeds all"
GoForm "seedsinput"
selectall
select OrderNo=Numb
if   info("Selected") < info("Records") 
stop
else
message "nothing found"
endif
;this only works if you have the up-to-date Seeds All file in your Tally folder. Up to date Sees All lives in Tyee˛/OfficeInfo/øˇ˛show pullers&checkersˇopenfile "pullers&checkers"
window "pullers&checkers"
Synchronize˛/show pullers&checkersˇ˛MarkPickupsˇIf info("selected") < info("records")
SelectAll
Endif
loop
getscrap "Next!"
Numb=val(clipboard())
find OrderNo=Numb
if  info("Found") AND Status≠"" AND ShipCode="P"
OrderPickedUp="X"
else
Message "Are you sure you know what you're doing?"
endif
until val(clipboard())=0
˛/MarkPickupsˇ˛set_check_minimumˇ;this flips the refund minimum to generate checks between $2 and zero
if zcheckmin=2
zcheckmin=0
else
zcheckmin=2
endif
message "The minimum refund to generate checks is now set to $"+str(zcheckmin)+"."˛/set_check_minimumˇ˛reset recall flagˇ;this flips the recall from N to Y or vice versa
if zrecall="N"
zrecall="Y"
else
zrecall="N"
endif
message "The recall flag has been set to "+zrecall+"."˛/reset recall flagˇ˛(extras)ˇ˛/(extras)ˇ˛clear_vmailˇvmail=""˛/clear_vmailˇ˛BulkCheckˇselectall
select BulkFlag≠"X" and «1SpareText» ≠ ""
GoForm "roberta's form"˛/BulkCheckˇ˛selectcoreˇSelect str(OrderNo) notcontains "."
˛/selectcoreˇ˛search_allˇglobal Numb
If info("selected") < info("records")
SelectAll
EndIf
getscrap "Next!"
Numb=clipboard()
if val(Numb)≠0 AND Numb NOTCONTAINS "-"
selectwithin OrderNo=val(Numb) OR Zip=val(Numb) OR Z=val(Numb) OR «C#»=val(Numb) OR MAd CONTAINS Numb OR SAd CONTAINS Numb
else
Selectwithin Con CONTAINS Numb OR Group CONTAINS Numb OR Telephone CONTAINS Numb
endif
if  info("Selected") = info("Records")
message "Nothing found."
endif˛/search_allˇ˛CheckLockedˇlocal zlocks, zserverID
message "This macro checks the record server id of the order you're looking at, then reports the id's of any locked records and attempts to unlock the order you're looking at, if you choose to."
zserverID=info("serverrecordid")
    if zserverID=0
    message "No id was reported? Go ask Ryan."
    stop
    endif
message "The server id of this record is "+str(zserverID)+" ."

lockedrecordlist zlocks
    if zlocks=""
    message "Nothing is locked."
    stop
    endif
message "These are the currently locked records:"+¶+zlocks

If zlocks contains str(zserverID)
YesNo "Do you want to try to unlock this order?"
    if clipboard()="Yes"
    forceunlockrecord    
    endif
else
message "This order is not a locked order. If you're still having trouble let Ryan know."
endif

YesNo "Do you want to double check?"
    if clipboard()="Yes"
lockedrecordlist zlocks
    if zlocks=""
    message "Nothing is locked."
    stop
    endif
If zlocks contains str(zserverID)
YesNo "Do you want to try to unlock this order again?"
    if clipboard()="Yes"
    forceunlockrecord    
    endif
else
message "This order is not a locked order. If you're still having trouble let Ryan know."
endif
    endif
    ˛/CheckLockedˇ˛find_lockedˇlocal zlockedord, zlocks
selectall

lockedrecordlist zlocks
    if zlocks=""
    message "Nothing is locked."
    ;stop
    else
    message "These are the currently locked records:"+¶+zlocks
    endif

    if zlocks≠""
GetScrapOK "what record id?"
if clipboard()≠""
find  info("ServerRecordID") = val(clipboard())
endif
    endif
˛/find_lockedˇ˛unlockˇforceunlockrecord˛/unlockˇ˛find_by_record_idˇGetScrapOK "what record id?"
if clipboard()≠""
find  info("ServerRecordID") = val(clipboard())
endif
˛/find_by_record_idˇ˛reset_to_original_orderˇYesNo "Reset order to office file?"
    if clipboard()="Yes"
Order=«Original Order»
        YesNo "Clear picksheet?"
        if clipboard()="Yes"
        PickSheet=""
        endif
            YesNo "Clear backorder?"
            if clipboard()="Yes"
            Backorder=""
            endif
   else
   Message "nothing was done"
   endif˛/reset_to_original_orderˇ˛clearpicksheetˇPickSheet=""˛/clearpicksheetˇ˛.clearOriginalPickˇ;OriginalPicksheet field is no longer in use
OriginalPicksheet=""˛/.clearOriginalPickˇ˛---ˇ˛/---ˇ˛exportPulledOrdersˇ;for convenience, this downloads the selected orders into the 'backorder' file.
waswindow= info("WindowName") 
form=info("FormName")
        if form="seedspagecheck"    
openfile "backorders"
oldwindow=info("windowname")
window waswindow
firstrecord
        back=""
            loop
        linenu=1
        loop
        back=back+extract(extract(Order,¶,linenu),¬,1)+¬+?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[1,"-"][1,-2],extract(extract(Order,¶,linenu),¬,2))+¬+
          ?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[-1,-1],"")+¬
          +extract(extract(Order,¶,linenu),¬,3)+¬+extract(extract(Order,¶,linenu),¬,4)[1,-1]+¬+extract(extract(Order,¶,linenu),¬,5)+¶
        linenu=linenu+1
        Until  extract(extract(Order,¶,linenu),¬,2)=""   
            downrecord
            until  info("Stopped") 
window oldwindow
openfile "&@back"
save
window waswindow
        endif˛/exportPulledOrdersˇ˛order_exporterˇ;this downloads the selected orders into the 'order_export' file with order numbers.
waswindow= info("WindowName") 
form=info("FormName")
        if form="seedspagecheck"    
openfile "order_export"
oldwindow=info("windowname")
window waswindow
firstrecord
        back=""
            loop
        linenu=1
        loop
        back=back+str(OrderNo)+¬+extract(extract(Order,¶,linenu),¬,1)+¬+?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[1,"-"][1,-2],extract(extract(Order,¶,linenu),¬,2))+¬+
          ?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[-1,-1],"")+¬
          +extract(extract(Order,¶,linenu),¬,3)+¬+extract(extract(Order,¶,linenu),¬,4)[1,-1]+¬+extract(extract(Order,¶,linenu),¬,5)+¶
        linenu=linenu+1
        Until  extract(extract(Order,¶,linenu),¬,2)=""   
            downrecord
            until  info("Stopped") 

window oldwindow
openfile "&@back"
save
window waswindow
        endif˛/order_exporterˇ˛ExportPrePulledˇ;for convenience, this downloads the selected orders into the 'backorder' file.
waswindow= info("WindowName") 
form=info("FormName")
        if form="seedspagecheck"    
openfile "backorders"
oldwindow=info("windowname")
window waswindow
firstrecord
        back=""
            loop
        linenu=1
        loop
        back=back+extract(extract(Order,¶,linenu),¬,1)+¬+?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[1,"-"][1,-2],extract(extract(Order,¶,linenu),¬,2))+¬+
          ?(extract(extract(Order,¶,linenu),¬,2) CONTAINS "-",extract(extract(Order,¶,linenu),¬,2)[-1,-1],extract(extract(Order,¶,linenu),¬,3))+¬
          +extract(extract(Order,¶,linenu),¬,4)[1,-1]+¶
        linenu=linenu+1
        Until  extract(extract(Order,¶,linenu),¬,2)=""   
            downrecord
            until  info("Stopped") 
window oldwindow
openfile "&@back"
save
window waswindow
        endif˛/ExportPrePulledˇ˛SingleItemExporterˇ;this looks for a specific item and exports that line of data related to that order to the order&line file. Way cool!
local ordernu, itemnu
backline=""
added=""
itemnu=""
firstrecord

;this piece gets, configures and selects for the item you're looking for
    GetScrapOK "What's the item number"
    if clipboard() contains "-"
        if length(clipboard()[1,"-"][1,-2])=3
        itemnu=" "+clipboard()
        else
        itemnu=clipboard()
        endif
    endif
    if clipboard() notcontains "-"
        if length(clipboard())=3
        itemnu=" "+clipboard()
        else
        itemnu=clipboard()
        endif
    endif
   ; message itemnu
select Order contains itemnu    
   ; stop
   
;this piece extracts the item from the orders
itemnu="*"+itemnu+"*"
linenu=1
ordernu=1
loop
linenu=arraysearch(Order,itemnu, linenu,¶)
stoploopif linenu=0
backline=str(OrderNo)+¬+Con+¬+City+¬+St+¬+str(Zip)+¬+Telephone+¬+extract(Order,¶,linenu)+¶
added=added+backline
linenu=linenu+1
until info("empty")

loop
downrecord
ordernu=ordernu+1
linenu=1

loop
linenu=arraysearch(Order,itemnu, linenu,¶)
stoploopif linenu=0
backline=str(OrderNo)+¬+Con+¬+City+¬+St+¬+str(Zip)+¬+Telephone+¬+extract(Order,¶,linenu)+¶
added=added+backline
linenu=linenu+1
until info("empty")

until  info("EOF") 

;message added+"-yo"

openfile "order&line"
openfile "&@added"˛/SingleItemExporterˇ˛export_all_addressedˇ;this exports all items for the orders selected with basic address info
;select around group cover sheets
waswindow=zyear+"seedstally:seedspagecheck"
openfile "order&line"
deleteall
window waswindow
backline=""
added=""
firstrecord

loop

ArrayFilter Order, backline, ¶, str(OrderNo)+¬+Con+¬+City+¬+St+¬+str(Zip)+¬+Telephone+¬+extract(Order,¶,seq())
    window "order&line:secret"
    openfile "+@backline"
    window waswindow
;added=added+backline
downrecord

until  info("stopped") 
˛/export_all_addressedˇ˛sub_exporterˇ;this exports any items where there have been substitutions made
local ordernu, itemnu
    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
backline=""
added=""
itemnu="sub-"
select Order contains "sub-"
firstrecord
   
;this piece extracts the item from the orders
itemnu="*"+itemnu+"*"
linenu=1
ordernu=1
loop
linenu=arraysearch(Order,itemnu, linenu,¶)
stoploopif linenu=0
backline=str(OrderNo)+¬+Con+¬+str(Zip)+¬+Telephone+¬+extract(Order,¶,linenu)+¶
added=added+backline
linenu=linenu+1
until info("empty")

;message added
;stop
loop
downrecord
ordernu=ordernu+1
linenu=1

loop
linenu=arraysearch(Order,itemnu, linenu,¶)
stoploopif linenu=0
backline=str(OrderNo)+¬+Con+¬+str(Zip)+¬+Telephone+¬+extract(Order,¶,linenu)+¶
added=added+backline
linenu=linenu+1
until info("empty")

until  info("EOF") 

;message added+"-yo"

openfile "sub_report"
openfile "&@added"
    endif˛/sub_exporterˇ˛backorder_exporterˇ;this macro exports all orders backorders, typically used at the end of the season for reporting
;names are added for Nikos as a separate lookup from comments once all the dupes are removed
;remember nikos wants a printed report
;message folderpath(dbinfo("folder",""))
;stop

clipboard()=""
    if folderpath(dbinfo("folder","")) contains "backorders"
    NoYes "Do you want to export backorders for these orders?"
    if clipboard()="No"
    stop
    endif
if clipboard()="Yes"    
local namer
gosheet ;for faster macro run time
waswindow= info("WindowName") 

;if the file name gets too long the backorders by order file name will be crap
namer=""
namer="backorders"+  info("DatabaseFileName")["0-9",-1][" ",-4] ;may need tweaking each year
;message namer
;stop

select Backorder≠""
openfile "backorders by order"
oldwindow=info("windowname")
window waswindow
firstrecord
        back=""
            loop
        linenu=1
        loop
        back=back+str(OrderNo)+¬+extract(extract(Backorder,¶,linenu),¬,1)+¬+?(extract(extract(Backorder,¶,linenu),¬,2) CONTAINS "-",extract(extract(Backorder,¶,linenu),¬,2)[1,"-"][1,-2],extract(extract(Backorder,¶,linenu),¬,2))+¬+
          ?(extract(extract(Backorder,¶,linenu),¬,2) CONTAINS "-",extract(extract(Backorder,¶,linenu),¬,2)[-1,-1],"")+¬
          +extract(extract(Backorder,¶,linenu),¬,3)+¬+¬+extract(extract(Backorder,¶,linenu),¬,5)[1,-1]+¬+datepattern(FirstFillDate,"mm/dd/yy")+¶
        linenu=linenu+1
        Until  extract(extract(Backorder,¶,linenu),¬,2)=""   
            downrecord
            until  info("Stopped") 
window oldwindow
openfile "&@back"
saveas namer
field DupeCode
formulafill str(OrderNo)+"-"+str(Line)
save
window waswindow
endif
endif˛/backorder_exporterˇ˛copy_backorder_to_textˇmessage "this macro puts a text report on the clipboard of all backorders for the selected orders suitable for pasting into a message."
local zbackorderarray
zbackorderarray=""
arrayselectedbuild zbackorderarray,¶,"",?(Backorder≠"",str(OrderNo)+" : "+¶+Backorder,"")
clipboard()=zbackorderarray 
;message clipboard()˛/copy_backorder_to_textˇ˛soldout_reportˇ;this exports any items that were out of stocked
local ordernu, itemnu
    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
backline=""
added=""
itemnu="sold out"
firstrecord
   
;this piece extracts the item from the orders
itemnu="*"+itemnu+"*"
linenu=1
ordernu=1
loop
linenu=arraysearch(Order,itemnu, linenu,¶)
stoploopif linenu=0
backline=str(OrderNo)+¬+Con+¬+str(Zip)+¬+Telephone+¬+extract(Order,¶,linenu)+¶
added=added+backline
linenu=linenu+1
until info("empty")

;message added
;stop
loop
downrecord
ordernu=ordernu+1
linenu=1

loop
linenu=arraysearch(Order,itemnu, linenu,¶)
stoploopif linenu=0
backline=str(OrderNo)+¬+Con+¬+str(Zip)+¬+Telephone+¬+extract(Order,¶,linenu)+¶
added=added+backline
linenu=linenu+1
until info("empty")

until  info("EOF") 

;message added+"-yo"

openfile "soldout_report"
openfile "&@added"
    endif˛/soldout_reportˇ˛export_contact_infoˇ;this exports contact information for selected orders
added=""

arrayselectedbuild added,¶,"",str(OrderNo)+¬+Group+¬+Con+¬+MAd+¬+City+¬+St+¬+
                    str(Pattern(Zip,"#####"))+¬+Telephone+¬+Email
;bigmessage added
;stop

openfile "contact_info"
openfile "&@added"
save˛/export_contact_infoˇ˛export_to_walkinˇ;this exports order's filled information to the clipboard, skipping sold out items,
;for the order your looking at formatted 
;for the walkin register, ready to be pasted in
added=""
linenu=1

loop
    added=added+?(val((extract(extract(Order,¶,linenu),¬,4)))≠0,extract(extract(Order,¶,linenu),¬,2)[1,"-"][1,-2]+"+"+
    extract(extract(Order,¶,linenu),¬,2)[-1,-1]+"+"+
    strip(extract(extract(Order,¶,linenu),¬,4))+¶,"")
        added=arrayreplacevalue(added,"A","1","+")
        added=arrayreplacevalue(added,"B","2","+")
        added=arrayreplacevalue(added,"C","3","+")
        added=arrayreplacevalue(added,"D","4","+")
        added=arrayreplacevalue(added,"E","5","+")
        added=arrayreplacevalue(added,"K","6","+")
        added=arrayreplacevalue(added,"L","6","+")
linenu=linenu+1
Until  extract(extract(Order,¶,linenu),¬,2)=""   

bigmessage added
;stop

clipboard()=added
˛/export_to_walkinˇ˛----ˇ˛/----ˇ˛CheckCardRecordˇlocal shortCC, addpd, chkzip
waswindow=info("windowname")
    if CreditCard≠""
Numb=OrderNo
shortCC=val(CreditCard[-4,-1])
addpd=«1stPayment»
;chkzip=Zip
openfile zyear+"Authorize.NetCardRecord"
selectall
;Select Zip=val(chkzip)
Select «Invoice Number»=Numb  
Selectadditional «Card Number» = shortCC
Selectadditional «Settled Amount»=addpd
if  info("Selected") =  info("Records") 
message "Nothing found."
endif
    else
    message "There's no credit card associated with this order?!?"
    endif˛/CheckCardRecordˇ˛find_last_orderˇ    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
field OrderNo
sortup
Select OrderNo<989999
LastRecord
ono=OrderNo
selectall
find OrderNo=ono˛/find_last_orderˇ˛Dunumˇzdun=0
    GetScrapOK "Minimum balance due to search for?"
    zdun=0-val(clipboard())
selectall
select str(OrderNo) notcontains "."
selectwithin «BalDue/Refund» ≤ zdun
selectwithin ShipCode notcontains "D"
selectwithin Notes1 notcontains "barter"
selectwithin Notes1 notcontains "ignore"
YesNo "avoided orders marked billed?"
if clipboard()="Yes"
selectwithin Notes1 notcontains "billed"
endif˛/Dunumˇ˛balance_due_arrayˇlocal zowees, zoweestotal
arrayselectedbuild zowees,¶,"", ¬+"Order "+str(OrderNo)+" - $"+str(abs(«BalDue/Refund»))
clipboard()=zowees
message zowees
arrayselectedbuild zoweestotal,¶,"",abs(«BalDue/Refund»)
zoweestotal=arraynumerictotal(zoweestotal,¶)
message zoweestotal
clipboard()="total selected balance dues - $"+str(zoweestotal)+¶+zowees
message clipboard()

˛/balance_due_arrayˇ˛diagnosticsˇif zdiagnose=""
message "searching for unreworked orders with an o needs work."
bigmessage "remember to check groups manually (manufactured groups, especially) for weirdnesses (VolDisc, Discount, Paid,1stPayment,1stTotal,1stRefund & BalDue/Refund should all be cleared)."
zdiagnose="Y"
message "this macro only syncs the first time in a session"
Synchronize
endif

;    if info("Selected") < info("Records") ;this if clause fails after running unreworked orders??
    selectall
;    endif
field OrderNo
sortup
Numb=""
;zlevel=6

case zlevel=0
    YesNo "Select unreworked orders?"
if clipboard()="Yes"
select «9SpareText» = ""
Selectwithin (Order contains " b " or Order contains ¬+"b ") and Order notcontains "B My Baby"
Selectadditional Order contains " o " or Order contains ¬+"o "
selectwithin Notes1 notcontains "all set"
zlevel=1
stop
endif

case zlevel=1
    YesNo "Select Order still contains backorder?"
if clipboard()="Yes"
select Order contains "backorder"
selectwithin Notes1 notcontains "all set"
zlevel=2
stop
endif

case zlevel=2
    YesNo "Select Status=""?"
if clipboard()="Yes"
select Status=""
selectwithin Notes1 notcontains "all set"
zlevel=3
stop
endif

case zlevel=3
    YesNo "Select Status≠Com?"
if clipboard()="Yes"
select Status="Com"
selectreverse
selectwithin Notes1 notcontains "all set"
zlevel=4
stop
endif

case zlevel=4
    YesNo "Select Status≠"" AND Paid=0?" ;typically purchase orders and side deals
if clipboard()="Yes"
select Status≠"" and Paid=0 and str(OrderNo) notcontains "." and OrderNo>val(Numb) and GrTotal≠0
selectwithin Notes1 notcontains "all set"
zlevel=5
stop
endif

case zlevel=5
    YesNo "Select «1stTotal»=0 or looks off?"
if clipboard()="Yes"
select «1stTotal»=0 and str(OrderNo) notcontains "."
selectadditional «1stTotal»≠0 and str(OrderNo) contains "."
selectwithin Notes1 notcontains "all set"
selectwithin Notes1 notcontains "order cancelled"
selectwithin OrderComments notcontains "empty order"
zlevel=6
stop
endif

case zlevel=6
    YesNo "Select AddPay2≠0?"
if clipboard()="Yes"
select AddPay2≠0 and OrderNo>val(Numb) and Order notcontains "add"
selectwithin Notes1 notcontains "all set"
zlevel=7
stop
endif

case zlevel=7
    YesNo "Select BalDue/Refund>0?" ;we owe them money
if clipboard()="Yes"
select «BalDue/Refund»>0
selectwithin Notes1 notcontains "all set"
selectwithin Status≠""
zlevel=8
stop
endif

case zlevel=8
    YesNo "Select patron+RealTax≠OrderTotal?" ;due to donations
if clipboard()="Yes"
select Status≠"" and Patronage+RealTax≠OrderTotal and OrderNo>val(Numb)
selectwithin Notes1 notcontains "all set"
zlevel=9
stop
endif

case zlevel=9
    YesNo "Select Pd≠1stPayment and Bal<0?" ;find quirks
if clipboard()="Yes"
select Status≠"" and Paid≠«1stPayment» and «BalDue/Refund»<0 and «BalDue/Refund»≠-.01 and OrderNo>val(Numb)
selectwithin Notes1 notcontains "all set"
zlevel=10
stop
endif

case zlevel=10
    YesNo "Select OrderTotal, Patronage or RealTax ≠ 0 for group pieces?"
if clipboard()="Yes"
select (str(OrderNo) contains "." and OrderTotal <> 0) OR (str(OrderNo) contains "." and Patronage <> 0) OR (str(OrderNo) contains "." and RealTax <> 0)    
selectwithin Notes1 notcontains "all set"
zlevel=11
stop
endif

case zlevel=11
    YesNo "Select OrderTotal+Donation+Membership≠GrTotal?" ;does not compute for group pieces as they have no grand total
if clipboard()="Yes"
select OrderTotal+Donation+Membership≠GrTotal and str(OrderNo) notcontains "."
selectwithin Notes1 notcontains "all set"
zlevel=12
stop
endif

case zlevel=12
    YesNo "Select order with no payments apparently? Typically meaning 1st payment field was not filled in at the office."
if clipboard()="Yes"
select «1stPayment»=0 and str(OrderNo) notcontains "." and (AddPay1+AddPay2+AddPay3+AddPay4+AddPay5+AddPay6)<Paid
selectwithin Notes1 notcontains "all set"
zlevel=13
stop
endif

case zlevel=13
    YesNo "Select RealTax≠SalesTax?"; does not compute for group pieces
if clipboard()="Yes"
select RealTax≠SalesTax+AddOnTax and str(OrderNo) notcontains "."
selectwithin Notes1 notcontains "all set"
zlevel=14
stop
endif

case zlevel=14
    YesNo "GrTotal≠«1stPayment»+all AddPay's-«Donated_Refund»-«BalDue/Refund»?"; looks for bad payment trail
if clipboard()="Yes"
select GrTotal≠«1stPayment»+AddPay1+AddPay2+AddPay3+AddPay4+AddPay5+AddPay6-«Donated_Refund»-«BalDue/Refund»
selectwithin Notes1 notcontains "all set"
zlevel=15
stop
endif

case zlevel=15
zlevel=0

endcase˛/diagnosticsˇ˛reset_zlevel_diagnosticˇzlevel=0˛/reset_zlevel_diagnosticˇ˛batch_subtotalˇlocal zsubtotalarray, zarraytotal
zsubtotalarray=""
zarraytotal=0
call selectcore
getscrapok "first order number"
selectwithin OrderNo ≥ val(clipboard())
getscrapok "last order number"
selectwithin OrderNo ≤ val(clipboard())
    yesno "is this the selection to total the subtotals of?"
    if clipboard()="Yes"
arrayselectedbuild zsubtotalarray,",","",Subtotal
;message zsubtotalarray
zarraytotal=arraynumerictotal(zsubtotalarray,",")
message zarraytotal
    endif
    ˛/batch_subtotalˇ˛.print coversheetˇwaswindow= info("WindowName") 
;GoForm "seedscoversheet"
PrintUsingForm "", "seedscoversheet"
PrintOneRecord
;goform waswindow[":",-1][2,-1]
˛/.print coversheetˇ˛UpdateDelinkedCommentsˇopenfile zcomyear+"SeedsComments linked"
ReSynchronize
save
closefile
openfile zcomyear+"SeedsComments"
openfile "&&"+zcomyear+"SeedsComments linked"
save
;message "I'm working on it."
window "Hide This Window"
window waswindow
˛/UpdateDelinkedCommentsˇ˛UnReWorkedˇ;this macro looks for orders that have not been reworked after initiating a manual status change, also done in diagnostics
local zfirstsort
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
field OrderNo
sortup
select «9SpareText» = ""
    zfirstsort = info("Selected") 
Selectwithin Order contains "b     " and Order notcontains "sub" and 
    Order notcontains "limit" and Order notcontains "pack" and Order notcontains "give"
     and Order notcontains "see RB"
    ;stop
Selectadditional Order contains "o     "
    if info("Selected") = zfirstsort
    message "Nothing found."
    selectall
    endif
˛/UnReWorkedˇ˛.GetGroupˇ    if folderpath(dbinfo("folder","")) CONTAINS "mailing"
ono=int(OrderNo) ;6-2 change
selectall
Select int(OrderNo)=ono ;6-2 change
    endif
˛/.GetGroupˇ˛getGroupˇono=int(OrderNo) ;6-2 change
SelectAll
Select int(OrderNo)=ono ;6-2 change
˛/getGroupˇ˛get trackingˇclipboard()=lookupall(zyear+"shipping","C#",«2SpareText»,"UPS#",¶)
message clipboard()˛/get trackingˇ˛writeemailˇif Email≠""
clipboard()="'"+Con+"' <"+Email+">"
message clipboard()
else
message "There's no email address. Call or write."
endif˛/writeemailˇ˛FindOtherOrdersˇlocal zcc
;zcc=«C#»  switched to email due to faster shipping turnaround 2-8-16
zcc=Email
selectall
select Email=zcc˛/FindOtherOrdersˇ˛sidecarˇ    if  info("Windows") notcontains "facilsidecar"
    WindowBox "30 40 580 950"
    openform "facilsidecar"
    else
    goform "facilsidecar"
    zoomwindow 30,40,550,920,""
    endif
˛/sidecarˇ˛MaineTaxTotalsˇSelect FillDate ≥ month1st(today()-30) and FillDate <month1st(today()) 
and Status contains "Com" and OrderNo = int(OrderNo)
gosheet
field OrderTotal
total
˛/MaineTaxTotalsˇ˛patronageUpdate/9ˇ;stop
loop
YesNo "Update patronage?"
    if clipboard()="Yes"
RealTax=SalesTax+AddOnTax
Patronage=OrderTotal-RealTax
downrecord
    endif
until  clipboard()="No"
˛/patronageUpdate/9ˇ˛huh/8ˇGrTotal=0
;if  (OrderNo>10000 AND OrderNo<30000) OR (OrderNo>70000 AND OrderNo<100000)
;if Pool≠2       ;pool 2 designation is for orders with a set amount for shipping
;«$Shipping»=0
;endif
;endif

    case OrderNo<30000 OR (OrderNo>70000 AND OrderNo<100000)

;for group pieces
                if OrderNo≠val(str(OrderNo)[1,5])
                    MemDisc=?(MemDisc>0, float(Subtotal-OGSTotal)*float(.01),0)
                    AdjTotal=Subtotal-MemDisc
                    SalesTax=0
;                    SalesTax=?(Taxable="Y", float(AdjTotal)*float(.055),0) ;sales tax is now determined by the group coordinators's tax status and delivery address
                    call ".salestax"
                    OrderTotal=0
                    Patronage=0
                    RealTax=0
                    Paid=0
                    «1stPayment»=0
                    «1stTotal»=0
                    «1stRefund»=0
                    «BalDue/Refund»=0
                    VolDisc=0
                    Discount=0
               Goto end
                endif
                
;now do regular orders and group cover sheets
        VolDisc=(float(Subtotal-OGSTotal)*float(Discount))
        
;David and I agreed to not cover the unlikely case that the entire group does not get the member discount 12-26-12
;meaning all groups with any member discount will get 1% at this point - sales tax will also be adjusted similarly
;                if vmemflag=0      ;does not calculate 1% member discount when group pieces are members individually and the group is not a member.
                MemDisc=?(MemDisc>0, float(Subtotal-OGSTotal)*float(.01),0)
;                endif             
    ;but avoid adjusting tax total on groups - it's adjusted further down                    
    ;If Order notcontains "0000" AND Order ≠ "" AND (Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderComments CONTAINS "empty order")
    ;AND (Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderNo>99000)
    ;TaxTotal=?(MemDisc>0, float(TaxTotal)*float(.99),TaxTotal)
    ;endif     
                organic_total=?(MemDisc>0,float(organic_subtotal)*(1-(Discount+.01)),float(organic_subtotal)*(1-Discount)) 
       
        AdjTotal=Subtotal-VolDisc-MemDisc

;why not just use adjusted total for sales tax and skip tax total?, answer group processing needs accumulative taxable total
;    if MemDisc>0
;        SalesTax=?(Taxable="Y", float(float(TaxTotal)*float((1-(Discount+.01))))*float(.055),0)
;    else
;        SalesTax=?(Taxable="Y", float(float(TaxTotal)*float((1-Discount)))*float(.055),0)
;    endif
;        SalesTax=?(Taxable="N", 0, SalesTax)
call ".salestax"

;zones 6-8 for now this flag is for medium flat rate box use
;If (Zip ≥ 00600 AND Zip ≤ 00999) OR (Zip ≥ 29800 AND Zip ≤ 37599) OR (Zip ≥ 38000 AND Zip ≤ 39949) OR 
;    (Zip ≥ 42000 AND Zip ≤ 42499) OR (Zip ≥ 47500 AND Zip ≤ 47899) OR (Zip ≥ 50000 AND Zip ≤ 52899) OR 
;    (Zip ≥ 53500 AND Zip ≤ 54099) OR (Zip ≥ 54400 AND Zip ≤ 54499) OR (Zip ≥ 54600 AND Zip ≤ 59999) OR 
;    (Zip ≥ 61000 AND Zip ≤ 99999)
;zone=6    
;endif
if zcanada="No"
call ".zone"
endif

zshipwt=float((«1SpareNumber»+«2SpareNumber»)/28.35) ;convert seed+packet weight to ounces
;message zshipwt
;now add in packaging
    if zcanada="Yes"
    zshipwt=?(zshipwt≤15 AND zshipwt>0, zshipwt+1,?(zshipwt≤64 AND zshipwt>0, zshipwt+5,?(zshipwt>240,zshipwt+16,zshipwt)))
    else
    zshipwt=?(zshipwt≤15 AND zshipwt>0, zshipwt+1,?(zshipwt>240,zshipwt+16,zshipwt))
    endif
3SpareNumber=round(zshipwt,.01)

            if ShippingWt>0 AND Pool≠2 AND zcanada="No" ;pool 2 allows a fixed amount for shipping
            call ".shipping"
            endif
            
            if zcanada="Yes" AND Pool≠2 ;pool 2 allows a fixed amount for shipping
            «$Shipping»=0
                if ShipCode ≠ "P"
            call ".shipping_canada"
                endif
            endif

     if Subtotal=0
     «$Shipping»=0
     endif   

        OrderTotal=AdjTotal+SalesTax+Surcharge+«$Shipping»+AddOns
        GrTotal=OrderTotal+Donation+Membership
        
     endcase
     vmemflag=0
    end:˛/huh/8ˇ˛huh2ˇ;this copies a list of all the fields to the clipboard

clipboard()=dbinfo("fields","")
message clipboard() ˛/huh2ˇ˛huh3ˇ˛/huh3ˇ˛yupˇlocal zemails
arrayselectedbuild zemails,";",""," "+Email ;for building email groups
;arrayselectedbuild zemails," ",""," "+Email  ; for one-up selections
message zemails
clipboard()=zemails˛/yupˇ˛SeedOriginˇ;has not been used in years
stop
each=""
origin=""
linenu=1
    if  info("Windows") notcontains "SEEDSPECS"
    opensecret "SEEDSPECS"
    Synchronize
    window waswindow
    endif
loop
each=extract(extract(Order,¶,linenu),¬,1)+" "+extract(extract(Order,¶,linenu),¬,2)[1,-3]+" - "+
            lookup("SEEDSPECS","cat #",val(extract(extract(Order,¶,linenu),¬,2)[1,-3]),"common","",0)+?(lookup("SEEDSPECS","cat #",val(extract(extract(Order,¶,linenu),¬,2)[1,-3]),"origin","",0)≠""," - ","")+
            lookup("SEEDSPECS","cat #",val(extract(extract(Order,¶,linenu),¬,2)[1,-3]),"origin","",0)
    if origin=""
    origin=each
    else
    origin=origin+¶+each
    endif
linenu=linenu+1
until linenu>arraysize(Order,¶)
message origin
goform "origins"˛/SeedOriginˇ˛dimensionsˇlocal getdim
getdim=info("WindowVariables") 
message info("WindowBox") ˛/dimensionsˇ˛exportogsˇwaswindow=info("windowname")
select OGSOrder≠""
order=OGSOrder+¶
loop
stoploopif info("eof")
downrecord
order=order+OGSOrder+¶
until info("stopped")
selectall
openfile "collationtemp"
openfile "&@order"
field Item
groupup
field Qty
total
RemoveDetail  "Data"
lastrecord
deleterecord
save
export zyear+"seedsexport.txt", str(Item)+¬+str(Qty)+¶
closewindow
window waswindow
selectall
˛/exportogsˇ˛seedsmonthlyˇlocal seedsselected
select OrderNo = int(OrderNo)
seedsselected=info("selected")
;selectwithin FillDate ≥ month1st((month1st(today())-1)) and FillDate < month1st(today()) 
;   and Status contains "com" 
selectwithin Status contains "com"
;if info("selected")=seedsselected
;beep
;stop
;endif
//selectwithin arraycontains(ztaxstates,TaxState,",")

arrayselectedbuild raya,¶,"", "Seeds"+¬+str(OrderNo)+¬+Taxable+¬+TaxState+¬+Cit+¬+pattern(Z,"#####")+¬+¬+str(TaxRate)+¬+str(StateRate)+¬+
str(CountyRate)+¬+str(CityRate)+¬+str(SpecialRate)+¬+str(«$Shipping»)+¬+str(AdjTotal)+¬+str(TaxedAmount)+¬+str(SalesTax)+¬+datepattern(FillDate,"Month YYYY")+¬+¬+
str(StateTax)+¬+str(CountyTax)+¬+str(CityTax)+¬+str(SpecialTax)+¬+str(OrderTotal)
clipboard()=raya
openfile "WayfairSalesTax"
openfile "+@raya"
˛/seedsmonthlyˇ˛charactersˇmessage length(Order)˛/charactersˇ˛.1648 refundˇstop
local delicataline
loop
delicataline=""
linenu=1
loop
linenu=arraysearch(Order,"*1648*", linenu,¶)
stoploopif linenu=0
until info("found")
Order=arraychange(Order,arraychange(extract(Order,¶,linenu),"o",5,¬),linenu,¶)
call reworkchanges/®

if str(OrderNo) notcontains "."
«BalDue/Refund»=Paid-GrTotal
call ".refund"
Paid=Paid-«BalDue/Refund»
«BalDue/Refund»=0
«25Total»=OrderTotal-RealTax
endif
downrecord

linenu=1
added=""
loop
backline=""
linenu=arraysearch(Order,"*1648*", linenu,¶)
stoploopif linenu=0
backline=extract(Order,¶,linenu)
    if backline≠""
    if added=""
    added=backline
    else
added=added+¶+backline
    endif
    endif
linenu=linenu+1
until info("empty")


if arraysize(added,¶)>1 or Paid≠GrTotal
YesNo "Good to go?"
    if clipboard()="No"
    stop
    endif
endif
until info("EOF")
˛/.1648 refundˇ˛.1539 recallˇstop
;local delicataline
loop
;delicataline=""
linenu=1

loop
linenu=arraysearch(Order,"*1539*", linenu,¶)
stoploopif linenu=0
until info("found")

Order=arraychange(Order,arraychange(extract(Order,¶,linenu),"o",5,¬),linenu,¶)
call reworkchanges/®

if str(OrderNo) notcontains "."
«BalDue/Refund»=Paid-GrTotal
call ".refund"
Paid=Paid-«BalDue/Refund»
«BalDue/Refund»=0
Patronage=OrderTotal-RealTax
endif

downrecord

linenu=1
added=""
loop
backline=""
linenu=arraysearch(Order,"*1539*", linenu,¶)
stoploopif linenu=0
backline=extract(Order,¶,linenu)
    if backline≠""
    if added=""
    added=backline
    else
added=added+¶+backline
    endif
    endif
linenu=linenu+1
until info("empty")


if arraysize(added,¶)>1; or Paid≠GrTotal
YesNo "Good to go?"
    if clipboard()="No"
    stop
    endif
endif
until info("stopped")
˛/.1539 recallˇ˛.483 recall selectorˇ;individuals with 483 subbed for 489
selectall
select Order contains "sub 483"
selectwithin str(OrderNo) notcontains "."
;selectwithin Order notcontains "sub 483"
selectwithin OrderNo≠72078 and OrderNo≠90868; previously adjusted
selectwithin OrderNo≠10576 and OrderNo≠11264 and OrderNo≠74080 and OrderNo≠78884; doubles individuals

stop

;individuals with 483
selectall
select Order contains " 483"
selectwithin str(OrderNo) notcontains "."
selectwithin Order notcontains "sub 483"
selectwithin OrderNo≠72078 and OrderNo≠90868; previously adjusted
selectwithin OrderNo≠10576 and OrderNo≠11264 and OrderNo≠74080 and OrderNo≠78884; doubles individuals
selectwithin OrderNo ≥11000 and OrderNo<12000

;groups with 483
selectall
select Order contains " 483"
selectwithin str(OrderNo) contains "."
selectwithin Order notcontains "sub 483"
selectwithin OrderNo≠72078 and OrderNo≠90868; previously adjusted
selectwithin OrderNo≠10576 and OrderNo≠11264 and OrderNo≠74080 and OrderNo≠78884; doubles individuals
selectwithin OrderNo≠99040.005 and OrderNo≠99046.001 ; previously adjusted groups

;groups with 483 subbed for 489
selectall
select Order contains "sub 483"
selectwithin str(OrderNo) contains "."
;selectwithin Order notcontains "sub 483"
selectwithin OrderNo≠72078 and OrderNo≠90868; previously adjusted
selectwithin OrderNo≠10576 and OrderNo≠11264 and OrderNo≠74080 and OrderNo≠78884; doubles individuals
selectwithin OrderNo≠99040.005 and OrderNo≠99046.001 ; previously adjusted groups

˛/.483 recall selectorˇ˛.483 recall generatorˇ;stop
;there's something about balance dues that are not factoring in correctly 6-20-19
;local delicataline
zcheckmin=0 ;this sets all refunds to generate checks
loop
;delicataline=""
linenu=1

loop
linenu=arraysearch(Order,"* 483*", linenu,¶)
stoploopif linenu=0
until info("found")

Order=arraychange(Order,arraychange(extract(Order,¶,linenu),"o",5,¬),linenu,¶)
call reworkchanges/®

if str(OrderNo) notcontains "."
«BalDue/Refund»=Paid-GrTotal
call ".refund"
Paid=Paid-«BalDue/Refund»
«BalDue/Refund»=0
Patronage=OrderTotal-RealTax
endif

downrecord

linenu=1
added=""
loop
backline=""
linenu=arraysearch(Order,"* 483*", linenu,¶)
stoploopif linenu=0
backline=extract(Order,¶,linenu)
    if backline≠""
    if added=""
    added=backline
    else
added=added+¶+backline
    endif
    endif
linenu=linenu+1
until info("empty")


if arraysize(added,¶)>1; or Paid≠GrTotal
YesNo "Good to go?"
    if clipboard()="No"
    stop
    endif
endif
until info("stopped")
˛/.483 recall generatorˇ˛.recall refund w addressˇzmethod="Cash"
data= info("DatabaseName") 
If «BalDue/Refund» ≥ zcheckmin OR (zcanada="Yes" AND «BalDue/Refund» > .10) OR vmail="pBm"
;    YesNo "Do you want to write a check for "+pattern(«BalDue/Refund», "$#.##")+"?"
 ;       if clipboard()="Yes"
        zmethod="Check"
;       endif


call .RefundRecord

if zmethod="Check"
        waswindow=info("windowname")
        window zyear+"RefundRegister"
        insertbelow
        Name=grabdata(data, Group)+" or "+grabdata(data, Con)
        Name=?(grabdata(data, Group)="", grabdata(data, Con),Name)
        «$Amount»=grabdata(data, «BalDue/Refund»)
         Amount=grabdata(data, «BalDue/Refund»)
       Date=today()
       OrderNO=grabdata(data, OrderNo)
       Con=grabdata(data, Con)
       Group=grabdata(data, Group)
       MAd=grabdata(data, MAd)
       City=grabdata(data, City)
       St=grabdata(data, St)
       Zip=grabdata(data, Zip)
        OpenForm "Check2"
        PrintOneRecord
        CloseWindow
        FirstRecord
        LastRecord
        Save
        window waswindow

        ;the Morse Pitts warning
        if «BalDue/Refund» ≥ 2000
        message "Please warn Gene about the size of this check so we're ready at the bank! Thanks."
        endif

       else 
       message "Don't forget to give a CASH REFUND of "+pattern(«BalDue/Refund», "$#.##")+"."

endif
else 
call .RefundRecord
message "Don't forget to give a CASH REFUND of "+pattern(«BalDue/Refund», "$#.##")+"."         
endif
˛/.recall refund w addressˇ˛.FindCCˇlocal zcode
zcode=«C#»
window zyear+" mailing list.warehouse"
select «C#»=val(zcode)
window zyear+"seedstally:seedspagecheck"˛/.FindCCˇ˛.ccInitializeStorageˇopenfile "credit charges"
    If  info("Records") =1 ;tally folder must be on desktop for script to work
    applescript |||
    tell application "Finder"
	activate
	try
		select file "creditimport" of folder "32tally.seeds"
		delete selection
	on error
		tell application "Panorama"
			activate
		end tell
	end try
	tell application "Panorama"
		activate
	end tell
end tell
    |||
    endif
˛/.ccInitializeStorageˇ˛.old_baldueˇlocal charge, data
data= info("DatabaseName") 
if CreditCard="" or CreditCard="0"
message "Don't forget the balance due of "+pattern(abs(«BalDue/Refund»), "$#.##")
Refund=0
goto bottom
endif
    if GrTotal<«1stTotal»
        if GrTotal+Refund≤«1stTotal»
        charge=GrTotal+Refund
        endif
            if GrTotal+Refund>«1stTotal»
            charge=«1stTotal»
                if Status="Com"
                Refund=«1stTotal»-GrTotal
                endif
            endif
    else
    charge=GrTotal
    if Status="Com"
    Refund=0
    endif
    endif
YesNo "Do you want to charge the balance due?"
if clipboard()="Yes"
window "credit charges"
If  info("Records") =1 ;tally folder must be on desktop for script to work
    applescript |||
    tell application "Finder"
	activate
	try
		select file "creditimport" of folder "32tally.seeds"
		delete selection
	on error
		tell application "Panorama"
			activate
		end tell
	end try
	tell application "Panorama"
		activate
	end tell
end tell
|||
endif
InsertBelow
«Credit Card #»=grabdata(data, CreditCard)
EXDate=grabdata(data, ExDate)
Amount=str(charge-grabdata(data, Paid))
Invoice=str(grabdata(data, OrderNo))
Zip=pattern(grabdata(data, Zip),"#####")
Save
window waswindow
    if Status="Com"
        if AddPay1=0
        AddPay1=charge-Paid
            else
            if AddPay1≠0 and AddPay2=0
            AddPay2=charge-Paid
            else
            AddPay3=AddPay3+charge-Paid
            endif
        endif
    Paid=GrTotal
    «BalDue/Refund»=0
    else
        if AddPay1=0
        AddPay1=charge-Paid
            else
            if AddPay1≠0 and AddPay2=0
            AddPay2=charge-Paid
            else
            AddPay3=AddPay3+charge-Paid
            endif
        endif
    Paid=Paid+charge
    endif
else
message "Don't forget the balance due of "+pattern(abs(«BalDue/Refund»),"$#.##")
endif
bottom:
˛/.old_baldueˇ˛.old_ManifestMailˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
If info("selected") < info("records")
SelectAll
EndIf
Synchronize
message "For accurate shipping dates, this needs to be run on the day the mail leaves the shop."

select MailFlag CONTAINS "p"
If info("selected") < info("records")
message "This manifest contains "+ str(info("selected")) +" piece(s)."
export "mail shipment"+" "+datepattern(today(),"mm/dd/yy"),Con+¬+Group+¬+MAd+¬+""+¬+City+¬+St+¬
        +str(Zip)+¬+str(OrderNo)+¬+"mail"+¬+str(«C#»)+¬+?(MailFlag contains "A","A",?(MailFlag contains "B","B",
        ?(MailFlag contains "C","C","")))+¬+datepattern(today(),"mm/dd/yy")+¶
arrayselectedbuild vupdate,¶,zyear+"seedstally", Con+¬+Group+¬+MAd+¬+""+¬+City+¬+St+¬
        +str(Zip)+¬+str(OrderNo)+¬+"mail"+¬+str(«C#»)+¬+?(MailFlag contains "A","A",?(MailFlag contains "B","B",
        ?(MailFlag contains "C","C","")))+¬+datepattern(today(),"mm/dd/yy")
openfile "33shipping"
openfile "+@vupdate"
window waswindow
field MailFlag
Fill ""
else
message "I can't find anything to ship?!?"
endif
    endif
selectall
select MailFlag contains "p"
if  info("Selected") <  info("Records") 
message "Mail flag failed to clear. The remaining p's will need to be cleared manually."
endif˛/.old_ManifestMailˇ˛.old_picksheetprintˇoldfile=  info("DatabaseName") 
waswindow=info("windowname")
form=info("FormName")
rayg=""
Synchronize
field OrderNo
SortUp
select Order contains "0000" OR Order = "" OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0"+¬+¬+¬+"0.00" AND OrderComments CONTAINS "empty order")
selectreverse
getscrap "first order to print"
Numb=val(clipboard())
ono=Numb
pickno=Numb
Selectwithin val(str(OrderNo)[1,5])≥Numb
getscrap "last order to print"
Numb=val(clipboard())
selectwithin val(str(OrderNo)[1,5])≤Numb
if form= "seedspagecheck" 
if OrderNo < 30000 OR (OrderNo>70000 AND OrderNo<100000)
selectwithin PickSheet=""
else
message "You're on the wrong form to be running these orders."
stop
endif        
endif
    if  info("Empty")
    message "Picksheets have already been printed for all the orders you selected."
    selectall
    stop
    endif
    
        if form="seedspagecheck"    
        openfile "32SeedsComments linked"
        ReSynchronize
        save
        closefile
        openfile "32SeedsComments"
        openfile "&&32SeedsComments linked"
        endif
save
window "Hide This Window"
window waswindow
call ".comments"

firstrecord
        if form="seedspagecheck"
        openform "seedspicksheet"
        endif
print ""
CloseWindow
    Selectwithin PickSheet CONTAINS "38)"
        If  info("Empty")
        Goto skip
    else
        if form="seedspagecheck"
        openform "seedspicksheet2"
        endif
        print ""
        CloseWindow
        window waswindow
   endif
    Selectwithin PickSheet CONTAINS "93)"
        If  info("Empty") 
        Goto skip
    else
        if form="seedspagecheck"
        openform "seedspicksheet3"
        endif
        print ""
        CloseWindow
        window waswindow
   endif
    Selectwithin PickSheet CONTAINS "148)"
        If  info("Empty") 
        Goto skip
    else
        if form="seedspagecheck"
        openform "seedspicksheet4"
        endif
        print ""
        CloseWindow
        window waswindow
    endif
    skip:
    
window waswindow
SelectAll
Select val(str(OrderNo)[1,5])≥pickno
selectwithin val(str(OrderNo)[1,5])≤Numb
zselect= info("Selected")   
        selectwithin OrderComments≠"" OR PickSheet CONTAINS "201)"
if info("Selected")<zselect
        if PickSheet CONTAINS "201)"
       message "Please remember to check comments on the orders currently selected. Bria thanks you!
Also, one or more order(s) has more than 200 items. Check with Bernice or Jen about this!"
       endif
        if PickSheet NOTCONTAINS "201)"
       message "Please remember to check comments on the orders currently selected. Bria thanks you!"
       endif
 else
 selectall
 endif      ˛/.old_picksheetprintˇ˛old_backorders_runˇzitem=""
brange=""
waswindow=info("WindowName") 
form= info("FormName")
selectall
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup
firstrecord

    openfile zcomyear+"SeedsComments linked"
    save
    closefile
    openfile zcomyear+"SeedsComments"
    openfile "&&"+zcomyear+"SeedsComments linked"
    save
    window "Hide This Window"
    window waswindow
    
NoYes "Did you forget to adjust Comments to account for items available only to backorders?"
    if clipboard()="Yes"
    stop
    endif
YesNo "Do you need to limit the range of orders?"
    If clipboard()="Yes"
    brange="Yes"
    else
    brange="No"
    endif
NoYes "Press - No - if you want to select or avoid specific items on backorder. Press - Yes - if you want to print all the backorders in the range, regardless of what's on them."
    if clipboard()="Yes"
        select Order contains "0000" OR Order = ""
        selectreverse
        selectwithin Status="B/O"
                if  info("Empty") 
                message "Nothing in this set has a Status of backorder."
                selectall
                stop
                endif
            if brange="Yes"
            getscrap "first order to print"
            Numb=val(clipboard())
            Selectwithin val(str(OrderNo)[1,5])≥Numb
            getscrap "last order to print"
            Numb=val(clipboard())
            selectwithin val(str(OrderNo)[1,5])≤Numb
            endif
    else
YesNo "Do you want to AVOID orders with specific backordered items (things not yet available)? Yes, lets you enter a list of the items or use the pre-built list of unavailable items."
    if clipboard()="Yes"
    YesNo "Do you want to use the pre-built list of unavailable items?"
    
                if clipboard()="Yes"
                Call "backorderlist"
                endif
                
         If clipboard()="No"       
        GetScrap "Enter first item #"
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        Select «Backorder» matchexact "*"+zitem+"*"
        Loop
        GetScrap "Enter next item # or zero"
        stoploopif val(clipboard())=0
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        SelectAdditional «Backorder» matchexact "*"+zitem+"*"
        Until  val(clipboard())=0
        endif
        
        SelectReverse
        selectwithin Status="B/O"
        selectwithin Order≠""

    ;this would add these items to a selection
    ;selectadditional (Backorder CONTAINS "3837" OR Backorder CONTAINS "2447" OR Backorder CONTAINS "2491") AND str(OrderNo) CONTAINS "."

        if  info("Empty") 
        message "Nothing in this set has a Status of backorder."
        selectall
        stop
        endif
            if brange="Yes"
            getscrap "first order to print"
            Numb=val(clipboard())
            Selectwithin val(str(OrderNo)[1,5])≥Numb
            getscrap "last order to print"
            Numb=val(clipboard())
            selectwithin val(str(OrderNo)[1,5])≤Numb
            endif
    else
YesNo "Do you want to print orders containing specific backordered items (things needing immediate attention)? Yes, lets you enter a list of the items."
    if clipboard()="Yes"        
        GetScrap "Enter first item #"
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        Select «Backorder» matchexact "*"+zitem+"*"
        Loop
        GetScrap "Enter next item # or zero"
        stoploopif val(clipboard())=0
            zitem=clipboard()
            if length(clipboard())=3
            zitem=" "+clipboard()
            endif
        SelectAdditional «Backorder» matchexact "*"+zitem+"*"
        Until  val(clipboard())=0
        selectwithin Status="B/O"
        selectwithin Order≠""
  
       ; this would allow us to skip group pieces or vice versa
       NoYes "Do you want to choose to run only individual or only group orders?"
       if clipboard()="Yes"
         NoYes "Do you want to select for individual orders only?"
         if clipboard()="Yes"
         selectwithin str(OrderNo) notcontains "."
         else
         NoYes "Do you want to select for group pieces only?"
         if clipboard()="Yes"
         selectwithin str(OrderNo) contains "."
         endif
         endif
       endif  
        
        if  info("Empty") 
        message "Nothing in this set has a Status of backorder."
        selectall
        stop
        endif
            if brange="Yes"
            getscrap "first order to print"
            Numb=val(clipboard())
            Selectwithin val(str(OrderNo)[1,5])≥Numb
            getscrap "last order to print"
            Numb=val(clipboard())
            selectwithin val(str(OrderNo)[1,5])≤Numb
            endif
    else
        stop
    endif
    endif
    endif
selectwithin Backorder CONTAINS ¬+"  0"+¬ 
zbackrunning="Yes"    
FirstRecord
Loop 
    If Backorder NOTCONTAINS ¬+"  0"+¬  ;this seems unnecessary-see above
    message "Order "+str(OrderNo)+"'s backorder was recently run and either wasn't processed or failed to update the server properly. It will not be run again. Check it out!"
    «10SpareText»="ran"
    endif
If Backorder CONTAINS ¬+"  0"+¬
Call ".backordercomments"
Call "reworkchanges/®"
endif
DownRecord
Until  info("Stopped") 
    SelectWithin «10SpareText»≠"ran"  ;this seems unnecessary-see above
    openform "BackorderPicksheet"
    print ""
    CloseWindow
if arraysize(Backorder,¶)>35
message "This order has more than 35 lines of backorder. Please check with Gene. They may not get all their backorders."
endif        
        done:
window waswindow
    YesNo "Do you want to export this set to generate a commodity list?"
    if clipboard()="Yes"
    call "exportbackorders"
    endif
firstrecord
brange=""
zitem=""
zbackrunning="No"    
˛/old_backorders_runˇ˛.peasˇ;remember to adjust these to update RealTax and Patronage
local pealine, peanum, peaaddon, pearefund, peanum2, pealine2, freerefund
data= info("DatabaseName") 
peanum=" 893"
pealine=arraysearch(Order,"*"+peanum+"*",1,¶)
peanum2=" 761"
pealine2=arraysearch(Order,"*"+peanum2+"*",1,¶)
peaaddon=0
pearefund=0
freerefund=0
if pealine≠0
;message array(Order,pealine,¶)
    if extract(extract(Order,¶,pealine),¬,5) contains "organic" OR
    extract(extract(Order,¶,pealine),¬,5) contains "here" OR
    extract(extract(Order,¶,pealine),¬,5) contains "sub 883"
Subtotal=Subtotal-val(extract(extract(Order,¶,pealine),¬,7))
TaxTotal=TaxTotal-val(extract(extract(Order,¶,pealine),¬,7))
pearefund=val(extract(extract(Order,¶,pealine),¬,7))
Order=arraychange(Order,extract(extract(Order,¶,pealine),¬,1)
            +¬+extract(extract(Order,¶,pealine),¬,2)
            +¬+extract(extract(Order,¶,pealine),¬,3)
            +¬+"  0"
            +¬+"out-of-stock   "
            +¬+extract(extract(Order,¶,pealine),¬,6)
            +¬+"   0.00"
            +¬+extract(extract(Order,¶,pealine),¬,8),pealine,¶)
     endif    
        
   if extract(extract(Order,¶,pealine),¬,5) contains "replacement" OR
    extract(extract(Order,¶,pealine),¬,5) contains "free"
    message "free"
    peaaddon=?(extract(extract(Order,¶,pealine),¬,2) contains "A","2.00",
                  ?(extract(extract(Order,¶,pealine),¬,2) contains "B","6.50",
                  ?(extract(extract(Order,¶,pealine),¬,2) contains "C","11.00","")))
      peaaddon=val(peaaddon)*val(extract(extract(Order,¶,pealine),¬,3))        
      peaaddon=peaaddon+?(Taxable="Y",peaaddon*.05,0)
      freerefund=pattern(peaaddon,"#.##")
      freerefund=val(freerefund)
      peaaddon=0-peaaddon           
            message peaaddon
            AddOns=AddOns+val(peaaddon)
   endif 
endif

if pealine2≠0
;message array(Order,pealine2,¶)
    if extract(extract(Order,¶,pealine2),¬,5) contains "organic" OR
    extract(extract(Order,¶,pealine2),¬,5) contains "here"
Subtotal=Subtotal-val(extract(extract(Order,¶,pealine2),¬,7))
TaxTotal=TaxTotal-val(extract(extract(Order,¶,pealine2),¬,7))
pearefund=pearefund+val(extract(extract(Order,¶,pealine2),¬,7))
Order=arraychange(Order,extract(extract(Order,¶,pealine2),¬,1)
            +¬+extract(extract(Order,¶,pealine2),¬,2)
            +¬+extract(extract(Order,¶,pealine2),¬,3)
            +¬+"  0"
            +¬+"out-of-stock   "
            +¬+extract(extract(Order,¶,pealine2),¬,6)
            +¬+"   0.00"
            +¬+extract(extract(Order,¶,pealine2),¬,8),pealine2,¶)
     endif    
endif

pearefund=pearefund+?(Taxable="Y",float(pearefund)*float(.05),0)
pearefund=pearefund-?(Discount>0,float(pearefund)*float(Discount),0)-
    ?(MemDisc>0,float(pearefund)*float(.01),0)
pearefund=pearefund+freerefund
    if «BalDue/Refund»<0 AND Status CONTAINS "Com"
    pearefund=pearefund+«BalDue/Refund»
    «BalDue/Refund»=0
    endif
pearefund=pattern(pearefund,"#.##")
    ;message pearefund

if pealine2≠0 Or pealine≠0
call .retotal
    if str(OrderNo) contains "."
    stop
    endif
Paid=Paid-val(pearefund)

    YesNo "Do you want to write a check for $"+pearefund+"?"
        if clipboard()="Yes"
        waswindow=info("windowname")
        window "33RefundRegister"
        insertbelow
        Name=grabdata(data, Group)+" or "+grabdata(data, Con)
        Name=?(grabdata(data, Group)="", grabdata(data, Con),Name)
        Con=grabdata(data, Con)
        Group=grabdata(data, Group)
        MAd=grabdata(data, MAd)
        City=grabdata(data, City)
        St=grabdata(data, St)
        Zip=grabdata(data, Zip)
        «$Amount»=val(pearefund)
         Amount=val(pearefund)
       Date=today()
       OrderNO=grabdata(data, OrderNo)
        OpenForm "PeaCheck"
        PrintOneRecord
        CloseWindow
        FirstRecord
        LastRecord
        Save
        window waswindow
        endif
endif

stop
linenu=1
zitem=""
zcomment=""
zline=1
loop
zitem=extract(extract(Backorder,¶,linenu),¬,2)
stoploopif zitem=""
zcomment=extract(extract(Backorder,¶,linenu),¬,4)[1,12]
zline=arraysearch(Order,"*"+zitem+"*",zline,¶)
stoploopif zline=0
Order=arraychange(Order,extract(extract(Backorder,¶,linenu),¬,1)
            +¬+extract(extract(Backorder,¶,linenu),¬,2)
            +¬+extract(extract(Backorder,¶,linenu),¬,3)
            +¬+extract(extract(Backorder,¶,linenu),¬,3)
            +¬+extract(extract(Backorder,¶,linenu),¬,4)
            +¬+extract(extract(Backorder,¶,linenu),¬,5)
            +¬+extract(extract(Backorder,¶,linenu),¬,6)
            +¬+extract(extract(Backorder,¶,linenu),¬,7),zline,¶)
linenu=linenu+1
zline=zline+1
until info("Empty")
˛/.peasˇ˛.pea selectˇSelect Order CONTAINS " 761-"
Selectwithin str(OrderNo) NOTCONTAINS "."
Selectwithin Status CONTAINS "B/O"˛/.pea selectˇ˛.CloseWindowˇ;the facilitators apparently don't want this, deactivated 2-19-18, remove entirely??
;       if folderpath(dbinfo("folder","")) CONTAINS "Facilitator"
;NoYes "If you close this window you won't be able to lookup Seed orders. Close it?"
;if clipboard()="Yes"
;closewindow
;endif
;        else
closewindow        
;        endif˛/.CloseWindowˇ˛.windowtobackˇWindowToBack zcomyear+"seedstally:facilitator view"˛/.windowtobackˇ˛check_200ˇlocal ordsz
ordsz= arraysize(Order,¶) > 200
message ordsz˛/check_200ˇ˛spare_text_keyˇ    if  info("Windows") notcontains "spare_text_key"
    WindowBox "30 640 380 1250"
    openform "spare_text_key"
    else
    goform "spare_text_key"
    zoomwindow 30,640,350,600,""
    endif
˛/spare_text_keyˇ˛copy_name_add_ph_emailˇclipboard()=Con+¶+MAd+¶+City+¶+St+¶+str(Zip)+¶+Telephone+¶+Email˛/copy_name_add_ph_emailˇ˛(canada)ˇ˛/(canada)ˇ˛SelectCanadiansˇselectall
    if timedifference(ztime,now())>ztimedif
Synchronize
ztime=now() ;resets sync timer
    endif
YesNo "Search for unprinted Canadian orders?"

If clipboard()="Yes"
select «9SpareText» ≠"" AND PickSheet="" AND Notes1 notcontains "cancelled" AND Order≠""
if  info("Selected") =  info("Records") 
message "nada"
endif
endif

If clipboard()="No"
    YesNo "Search for printed Canadian orders?"

If clipboard()="Yes"
select «9SpareText» ≠"" AND ((PickSheet≠"" AND Notes1 notcontains "cancelled") OR (PickSheet="" AND Subtotal≠0))
field OrderNo
sortup
if  info("Selected") =  info("Records") 
message "nada"
endif
endif
endif

If clipboard()="No"
    YesNo "Search for all Canadian orders?"

If clipboard()="Yes"
select «9SpareText» ≠""
field OrderNo
sortup
if  info("Selected") =  info("Records") 
message "nada"
endif
endif
endif

˛/SelectCanadiansˇ˛show_canada_sidecarˇ    if  info("Windows") notcontains "canadian_sidecar"
    WindowBox "30 640 380 950"
    openform "canadian_sidecar"
    else
    goform "canadian_sidecar"
    zoomwindow 30,640,350,300,""
    endif
˛/show_canada_sidecarˇ˛need_postal_infoˇselectall
YesNo "Search for Canadian orders needing post office weights and prices?"

If clipboard()="Yes"
select «9SpareText» ≠"" AND «4SpareNumber»=0 AND Notes1 notcontains "cancelled" AND PickSheet≠""
if  info("Selected") =  info("Records") 
message "nada"
endif
endif
˛/need_postal_infoˇ˛print_canada_batchˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
oldfile=  info("DatabaseName") 
waswindow=info("windowname")
form=info("FormName")
rayg=""
YesNo "Do you want to run picksheets for all the selected Canadian orders?"
    if clipboard()="Yes"
selectwithin ShipCode≠"D" OR (Order≠"" AND Order NOTCONTAINS "1)"+¬+"0")
    if ShipCode="D" OR Order="" OR Order CONTAINS "1)"+¬+"0"
    message "There's nothing here to print. All these orders are empty or declines."
    stop
    endif
        if form="seedspagecheck"    
        openfile zcomyear+"SeedsComments linked"
        ReSynchronize
        save
        closefile
        openfile zcomyear+"SeedsComments"
        openfile "&&"+zcomyear+"SeedsComments linked"
        endif
save
window "Hide This Window"
window waswindow
call ".comments"
selectwithin PickSheet≠"" ;this works primarily with the process in the totaller to avoid orders with items marked -here-.
if PickSheet≠""
firstrecord
if zcanadaorders=""
        PrintUsingForm "", "seedspicksheet"
        print ""
      Selectwithin PickSheet CONTAINS "41)"
    If  info("Empty")
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet2"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "96)"
    If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet3"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "151)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet4"
        print ""
    endif
endif

if zcanadaorders≠""
    if str(Zip) contains "0"
    field Zip
    ClearCell
    endif
;        PrintUsingForm "", "seedscoversheet"
;        print ""
        PrintUsingForm "", "canada_picksheet"
        print ""
      Selectwithin PickSheet CONTAINS "39)"
    If  info("Empty")
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet2"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "83)"
    If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet3"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "127)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet4"
        print ""
    endif
      Selectwithin PickSheet CONTAINS "171)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet5"
        print ""
    endif
endif
    skip:
    
       SelectWithin OrderComments≠"" OR PickSheet CONTAINS "201)" OR Notes1≠""
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!
Also, one or more order(s) may more than 200 items. Check with Bernice or Molly about this!"
else
message "Nothing to print."
endif
endif
    endif
zcanada="No"       
    ˛/print_canada_batchˇ˛.PicksheetPrint_forCANˇ;this macro was built to alternate printing modes for US and Canadian order, but proved impractical printer-wise
oldfile=  info("DatabaseName") 
waswindow=info("windowname")
form=info("FormName")
rayg=""
zlarge=0
Synchronize
ztime=now() ;resets sync timer
field OrderNo
SortUp
select Order contains "0000" OR Order = "" OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderComments CONTAINS "empty order")
    OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderNo>99000)
selectreverse
selectwithin ShipCode≠"D"

getscrap "first order to print"
Numb=val(clipboard())
ono=Numb
pickno=Numb
Selectwithin int(OrderNo)≥Numb ;6-2 change
getscrap "last order to print"
Numb=val(clipboard())
selectwithin int(OrderNo)≤Numb ;6-2 change

if form= "seedspagecheck" 
if OrderNo < 30000 OR (OrderNo>70000 AND OrderNo<1000000)
selectwithin PickSheet=""
else
message "You're on the wrong form to be running these orders."
stop
endif        
endif
    if  info("Empty")
    message "Picksheets have already been printed for all the orders you selected."
    selectall
    stop
    endif
    
        if form="seedspagecheck"    
        openfile zcomyear+"SeedsComments linked"
        ReSynchronize
        save
        closefile
        openfile zcomyear+"SeedsComments"
        openfile "&&"+zcomyear+"SeedsComments linked"
        endif
save
window "Hide This Window"
window waswindow
call ".comments"

firstrecord
zselect= info("Selected")   
    if zcanadaorders≠"" AND arraysize(zcanadaorders,",")<zselect
selectwithin «9SpareText»=""
if info("Selected")<zselect
    endif
        PrintUsingForm "", "seedspicksheet"
        print ""
      Selectwithin PickSheet CONTAINS "41)"
    If  info("Empty")
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet2"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "96)"
    If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet3"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "151)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet4"
        print ""
    endif
endif
    skip:

if zcanadaorders≠""
SelectAll
linenu=1
zprintrun=arraysize(zcanadaorders,",")
message "Remember to load Canadian picksheets."
loop
stoploopif zprintrun=0
ono=val(extract(zcanadaorders,",",linenu))
find OrderNo=ono
stoploopif linenu=zprintrun+1 OR «9SpareText»=""
        if  info("Found")
;message ono
            if Zip = 0 and «9SpareText» ≠ ""
            zcanada="Yes"
            endif
    If zcanada="Yes"
        if str(Zip) contains "0"
        field Zip
        ClearCell
        endif
PrintUsingForm "", "seedscoversheet"
print ""
PrintUsingForm "", "canada_picksheet"
PrintOneRecord
if PickSheet CONTAINS "41)"
PrintUsingForm "", "canada_picksheet2"
PrintOneRecord
endif
if PickSheet CONTAINS "85)"
PrintUsingForm "", "canada_picksheet3"
PrintOneRecord
endif
if PickSheet CONTAINS "129)"
PrintUsingForm "", "canada_picksheet4"
PrintOneRecord
endif
if PickSheet CONTAINS "173)"
PrintUsingForm "", "canada_picksheet5"
PrintOneRecord
endif
    endif
        endif
linenu=linenu+1
zcanada="No"
until info("Empty")
endif
    
SelectAll
Select int(OrderNo)≥pickno ;6-2 change
selectwithin int(OrderNo)≤Numb ;6-2 change
        selectwithin OrderComments≠"" OR PickSheet CONTAINS "201)" OR Notes1≠""
if info("Selected")<zselect
        if zlarge>1
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!
Also, one or more order(s) has more than 200 items. Check with Bernice or Molly about this!"
       endif
        if zlarge=0
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!"
       endif
 else
 selectall
 endif      
zcanada="No"       
zcanadaorders=""˛/.PicksheetPrint_forCANˇ˛.PrintThisPicksheet_forCAˇ;this macro was intended to print either US or Canadian orders, but proved impractical printer-wise
    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
zlarge=0
ono=OrderNo
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup
Select OrderNo=ono
oldfile=  info("DatabaseName") 
waswindow=info("windowname")
form=info("FormName")
rayg=""
if  info("Selected") =1
    if ShipCode="D"
    message "This order is being held in the office for lack of payment. It cannot be printed."
    stop
    endif
If Order contains "0000" OR Order = "" OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderComments CONTAINS "empty order")
    OR (Order CONTAINS "1)"+¬+"0"+¬+¬+"0" AND OrderNo>99000) OR PickSheet≠""
    message "There's nothing here to print or the picksheet has already been printed for this order or this is a group cover sheet."
    selectall
    Find OrderNo=ono
    stop
endif
        if form="seedspagecheck"    
        openfile zcomyear+"SeedsComments linked"
        ReSynchronize
        save
        closefile
        openfile zcomyear+"SeedsComments"
        openfile "&&"+zcomyear+"SeedsComments linked"
        endif
save
window "Hide This Window"
window waswindow
call ".comments"

firstrecord
if zcanadaorders=""
        PrintUsingForm "", "seedspicksheet"
        print ""
      Selectwithin PickSheet CONTAINS "41)"
    If  info("Empty")
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet2"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "96)"
    If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet3"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "151)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "seedspicksheet4"
        print ""
    endif
endif

if zcanadaorders≠""
    if str(Zip) contains "0"
    field Zip
    ClearCell
    endif
        PrintUsingForm "", "seedscoversheet"
        print ""
        PrintUsingForm "", "canada_picksheet"
        print ""
      Selectwithin PickSheet CONTAINS "41)"
    If  info("Empty")
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet2"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "85)"
    If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet3"
        print ""
   endif
      Selectwithin PickSheet CONTAINS "129)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet4"
        print ""
    endif
      Selectwithin PickSheet CONTAINS "173)"
        If  info("Empty") 
        Goto skip
    else
        PrintUsingForm "", "canada_picksheet5"
        print ""
    endif
endif
    skip:
    
    selectall
    Find OrderNo=ono
        if OrderComments≠"" OR PickSheet CONTAINS "201)" OR Notes1≠""
        if PickSheet CONTAINS "201)"
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!
Also, one or more order(s) has more than 200 items. Check with Bernice or Molly about this!"
       endif
        if PickSheet NOTCONTAINS "201)"
       message "Please remember to check comments and order notes on the orders currently selected. Bria thanks you!"
       endif
       endif
endif
    endif
zcanada="No"       
    ˛/.PrintThisPicksheet_forCAˇ˛run_canadian_boˇ    if folderpath(dbinfo("folder","")) CONTAINS "tally.seeds"
            if Zip = 0 and «9SpareText» ≠ ""
            NoYes "This is a Canadian order. Are you sure you want to run this?"
            if clipboard()="No"
            stop
            endif
            endif
    if str(OrderNo) contains "." OR Order=""
    NoYes "Do you want to run all the backorders for this group?"
    if clipboard()="Yes"
    Call "RunGroupBackorders"
    stop
    endif
    endif
YesNo "Do you want to run the backorder for this order?"
if clipboard()="Yes"    
ono=OrderNo
Synchronize
ztime=now() ;resets sync timer
field OrderNo
sortup
find OrderNo=ono
waswindow=info("WindowName") 
    if Backorder=""
    message "There is no backorder here to run!"
    stop
    endif
    If Backorder NOTCONTAINS ¬+"  0"+¬
    message "This backorder was recently run and either wasn't processed on the computer or failed to update the server properly. This process will now stop."
    stop
    endif
    
openfile zcomyear+"SeedsComments linked"
save
closefile
openfile zcomyear+"SeedsComments"
openfile "&&"+zcomyear+"SeedsComments linked"
save
window "Hide This Window"
window waswindow

    brange="No"
    zbackrunning="Yes"
Select OrderNo=ono
Call ".backordercomments"
Call "reworkchanges/®"
    openform "canada_picksheet"
    print ""
    CloseWindow
if arraysize(Backorder,¶)>35
message "This order has more than 35 lines of backorder. Please check with Ryan. They may not get all their backorders."
endif        
selectall
Find OrderNo=ono
brange=""
zbackrunning="No"
endif
    endif
˛/run_canadian_boˇ˛change_shippingˇgetscrapok "what should shipping be?"
«$Shipping»=val(clipboard())˛/change_shippingˇ˛field_listˇmessage visiblefields()
clipboard()=visiblefields()˛/field_listˇ˛SourceExportˇlocal Dictionary, ProcedureList
//this saves your procedures into a variable 
//step one
saveallprocedures "", Dictionary
clipboard()=Dictionary
//now you can paste those into a text editor and make your changes
STOP
 
//step 2
//this lets you load your changes back in from an editor and put them in
Dictionary=clipboard()
loadallprocedures Dictionary,ProcedureList
message ProcedureList //messages which procedures got changed˛/SourceExportˇ