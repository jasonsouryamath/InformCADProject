Attribute VB_Name = "Module1"
Public Sub main()
 On Error Resume Next
     Set oFactory = CreateObject("MSOS.clsObjectFactory")
       Set gOleServer = oFactory.CreateCADServer()
        Set gOleServer1 = oFactory.CreateCADServer()
       Dim id
       Dim obj0, obj1, obj2, obj3, obj4, obj5, obj6, obj7, obj8, obj9, obj10, obj11, obj12, obj13, obj14, obj15, obj16, obj17, obj18, obj19, obj20, obj21, obj22, obj23, obj24, obj25, obj26, obj27, obj28, obj29, obj30, obj31, obj32, obj33, obj35, obj36, obj37, obj38, obj39, obj40, obj41, obj42, obj43, obj44, obj45, obj46, obj47
        Dim obj1Resp0, obj1Resp1, obj1Resp2, obj1Resp3, obj1Resp4, obj1Resp5, obj1Resp6, obj1Resp7, obj1Resp8, obj1Resp9, obj1Resp10, obj1Resp11, obj1Resp12, obj1Resp13, obj1Resp14, obj1Resp15, obj1Resp16, obj1Resp17, obj1Resp18, obj1Resp19, obj1Resp20, obj1Resp21, obj1Resp22, obj1Resp23, obj1Resp24, obj1Resp25, obj1Resp26, obj1Resp27, obj1Resp28, obj1Resp29, obj1Resp30, obj1Resp31, obj1Resp32, obj1Resp33, obj1Resp34, obj1Resp35, obj1Resp36, obj1Resp37, obj1Resp38, obj1Resp39, obj1Resp40, obj1Resp41, obj1Resp42, obj1Resp43, obj1Resp44, obj1Resp45, obj1Resp46, obj1Resp47, obj1Resp48, obj1Resp49, obj1Resp50
         Dim obj1Resp51, obj1Resp52, obj1Resp53, obj1Resp54, obj1Resp55, obj1Resp56, obj1Resp57, obj1Resp58, obj1Resp59, obj1Resp60
         Dim obj34, obj49, obj48
         Dim obj1Resp551, obj1Resp552, obj1Resp553, obj1Resp554, obj1Resp555, obj1Resp556, obj1Resp557, obj1Resp558, obj1Resp559, obj1Resp560
         parametername = Agency_Type
           ' ServerRespId = gOleServer.GetAllActiveResponses1()
            ServerRespId1 = gOleServer.GetAllActiveResponses1(obj0, obj1, obj2, obj3, obj4, obj5, obj6, obj7, obj8, obj9, obj10, obj11, obj12, obj13, obj14, obj15, obj16, obj17, obj18, obj19, obj20, obj21, obj22, obj23, obj24, obj25, obj26, obj27, obj28, obj29, obj30, obj31, obj32, obj33, obj34, obj35, obj36, obj37, obj38, obj39, obj40, obj41, obj42, obj43, obj44, obj45, obj46, obj47, obj48, obj49)
           ' ServerRespId1 = gOleServer.GetAllActiveResponses1(Confirmation_Number:=id)

           ' '(parametername:=id)
             ServerRespId2 = gOleServer1.GetAllActiveResponses2(obj1Resp0, obj1Resp1, obj1Resp2, obj1Resp3, obj1Resp4, obj1Resp5, obj1Resp6, obj1Resp7, obj1Resp8, obj1Resp9, obj1Resp10, obj1Resp11, obj1Resp12, obj1Resp13, obj1Resp14, obj1Resp15, obj1Resp16, obj1Resp17, obj1Resp18, obj1Resp19, obj1Resp20, obj1Resp21, obj1Resp22, obj1Resp23, obj1Resp24, obj1Resp25, obj1Resp26, obj1Resp27, obj1Resp28, obj1Resp29, obj1Resp30, obj1Resp31, obj1Resp32, obj1Resp33, obj1Resp34, obj1Resp35, obj1Resp36, obj1Resp37, obj1Resp38, obj1Resp39, obj1Resp40, obj1Resp41, obj1Resp42, obj1Resp43, obj1Resp44, obj1Resp45, obj1Resp46, obj1Resp47, obj1Resp48, obj1Resp49, obj1Resp50, obj1Resp51, obj1Resp52, obj1Resp53, obj1Resp54, obj1Resp55, obj1Resp56, obj1Resp57, obj1Resp58) ', obj1Resp59)
            ServerRespId3 = gOleServer.GetAllActiveResponses3(obj1Resp551, obj1Resp552, obj1Resp553, obj1Resp554, obj1Resp555, obj1Resp556, obj1Resp557, obj1Resp558, obj1Resp559)
            

          ServerRespId = gOleServer.GetAllActiveResponses2(Late_Flag:=id)



 
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(App.Path + "\Cachedata.txt", True)
  
mer = "id|Master_Incident_Number|Response_Date|Agency_Type|Jurisdiction|Division|Battalion|Station|Response_Area|Response_Time_Criteria|Response_Plan|Incident_Type|Problem|Priority_Number|Priority_Description|Location_Name|Address|Apt_Nmbr|City|STATE|Postal_Code|County|Location_Type|Longitude|Latitude|Map_Info|Cross_Street|MethodOfCallRcvd|Call_Back_Phone|CallTaking_Performed_By|CallClosing_Performed_By|CallDisposition_Performed_By|OnCall_ID|Transfer_Return_Flag|System_Status_Level|ProQA_CaseNumber|Transfer_Table_ID|SentToBilling_Flag|ANI_Number|ALI_Info|Command_Channel|Primary_TAC_Channel|Alternate_TAC_Channel|Call_Disposition|Cancel_Reason|WhichQueue|Confirmation_Number|Base_Response_Number|Call_Is_Active|Call_Transfer_Receiving_Center|Caller_Type|Caller_Name|Caller_Location_Name|Caller_Address|Caller_Apt_Nmbr|Caller_City|Caller_State|Caller_Postal_Code|Caller_County|Caller_Location_Phone|Time_PhonePickUp|Time_FirstCallTakingKeystroke|"
st = "Time_CallEnteredQueue|Time_CallTakingComplete|Time_Incident_Under_Control"
ple = "|Time_CallClosed|Time_SentToOtherCAD|Time_First_Unit_Assigned|Time_First_Unit_Enroute|Time_First_Unit_Arrived|Elapsed_CallRcvd2InQueue|Elapsed_CallRcvd2CalTakDone|Elapsed_InQueue_2_FirstAssign|Elapsed_CallRcvd2FirstAssign|Elapsed_Assigned2FirstEnroute|Elapsed_Enroute2FirstAtScene|Elapsed_CallRcvd2CallClosed|Fixed_Time_PhonePickUp|Fixed_Time_CallEnteredQueue|Fixed_Time_CallTakingComplete|Fixed_Time_CallClosed|Fixed_Time_SentToOtherCAD|IncidentID|Fire_GeoArea|Police_GeoArea|AgencyType4_GeoArea|AgencyType5_GeoArea|AgencyType6_GeoArea|AgencyType7_GeoArea|AgencyType8_GeoArea|AgencyType9_GeoArea|AgencyType10_GeoArea|Late_Flag|CreatedByPrescheduleModule|PremiseID|Street_ID|Certification_Level|CurrentDivisionID|Stacked|BldngNmbr|Caller_Bldng_Nmbr|LateIcon|RequestToCancelIcon|UnreadIncident|UnreadComment|CautionNotes|Hazmat|PremiseHistory|NonGeoVerified|AttachmentIcon|HomeSectorID|CurrentSectorID|PriorityCode|ResponseTimeLate|RespReconfigState|"
tes = "NeedAddressUpdate|Notes|IncidentID"
 a.writeline (mer & st & ple & tes)
    On Error Resume Next
    For ai = 0 To (ServerRespId1)
         ab = obj0(ai) & "|" & obj1(ai) & "|" & obj2(ai) & "|" & obj3(ai) & "|" & obj4(ai) & "|" & obj5(ai) & "|" & obj6(ai) & "|" & obj7(ai) & "|" & obj8(ai) & "|" & obj9(ai) & "|" & obj10(ai) & "|" & obj11(ai) & "|" & obj12(ai) & "|" & obj13(ai) & "|" & obj14(ai) & "|" & obj15(ai) & "|" & obj16(ai) & "|" & obj17(ai) & "|" & obj18(ai) & "|" & obj19(ai) & "|" & obj20(ai) & "|" & obj21(ai) & "|" & obj22(ai) & "|"
         GH = obj23(ai) & "|" & obj24(ai) & "|" & obj25(ai) & "|" & obj26(ai) & "|" & obj27(ai) & "|" & obj28(ai) & "|" & obj29(ai) & "|" & obj30(ai) & "|" & obj31(ai) & "|" & obj32(ai) & "|" & obj33(ai) & "|" & "|" & obj35(ai) & "|" & "|" & obj37(ai) & "|" & obj38(ai) & "|" & obj39(ai) & "|" & obj40(ai) & "|" & obj41(ai) & "|" & obj42(ai) & "|" & obj43(ai) & "|" & obj44(ai) & "|" & obj45(ai) & "|" & obj46(ai) & "|" & obj47(ai) & "|" & obj48(ai) & "|" & obj49(ai) & "|"

         bc = obj1Resp0(ai) & "|" & obj1Resp1(ai) & "|" & obj1Resp2(ai) & "|" & obj1Resp3(ai) & "|" & obj1Resp4(ai) & "|" & obj1Resp5(ai) & "|" & obj1Resp6(ai) & "|" & obj1Resp7(ai) & "|" & obj1Resp8(ai) & "|" & obj1Resp9(ai) & "|" & obj1Resp10(ai) & "|" & obj1Resp11(ai) & "|" & obj1Resp12(ai) & "|" & obj1Resp13(ai) & "|" & obj1Resp14(ai) & "|" & obj1Resp15(ai) & "|" & obj1Resp16(ai) & "|" & obj1Resp17(ai) & "|" & obj1Resp18(ai) & "|" & obj1Resp19(ai) & "|" & obj1Resp20(ai) & "|" & obj1Resp21(ai) & "|" & obj1Resp22(ai) & "|" & obj1Resp23(ai) & "|" & obj1Resp24(ai) & "|" & obj1Resp25(ai) & "|" & obj1Resp26(ai) & "|" & obj1Resp27(ai) & "|" & obj1Resp28(ai) & "|" & obj1Resp29(ai) & "|" & obj1Resp30(ai) & "|" & obj1Resp31(ai) & "|" & obj1Resp32(ai) & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|" & "|"
        cd = obj1Resp42(ai) & "|" & obj1Resp43(ai) & "|" & obj1Resp44(ai) & "|" & obj1Resp45(ai) & "|" & obj1Resp46(ai) & "|" & obj1Resp47(ai) & "|" & obj1Resp48(ai) & "|" & obj1Resp49(ai) & "|" & obj1Resp50(ai) & "|" & obj1Resp51(ai) & "|" & obj1Resp52(ai) & "|" & obj1Resp53(ai) & "|" & obj1Resp54(ai) & "|" & obj1Resp55(ai) & "|" & obj1Resp56(ai) & "|" & obj1Resp57(ai) & "|" & obj1Resp58(ai) & "|"
         ef = obj1Resp551(ai) & "|" & obj1Resp552(ai) & "|" & obj1Resp553(ai) & "|" & obj1Resp554(ai) & "|" & obj1Resp555(ai) & "|" & obj1Resp556(ai) & "|" & obj1Resp557(ai) & "|" & obj1Resp558(ai) & "|" & obj1Resp559(ai)
             a.writeline (ab & GH & bc & cd & ef)
    Next
   
   




    a.Close

         


     



End Sub




