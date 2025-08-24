using XLConvert
using DataFrames

using Dates

function if_multiple(dividend, divisor, value)
xl_compare(xl_mod(dividend, divisor), 0) ? value : 0.0
end

function calculate_NA_Cost_Model_L_PSH_Q15(tab_Cost_Model_L_PSH_Q20_R25, NA_Cost_Model_L_PSH_O15, Surge_Facilities_Cost_Model_L_PSH_G15, Mean_Gross_Head__Cost_Model_L_PSH_C89, Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15)
	# =C15/C89
	L_per_H__Cost_Model_L_PSH_C95 = Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15 / Mean_Gross_Head__Cost_Model_L_PSH_C89 # Cost Model L-PSH C95
	@assert xl_compare(L_per_H__Cost_Model_L_PSH_C95, 10.855203619909503) # "Cost Model L-PSH!C95"
	# =IF(C95>=7,"Yes","No")
	Surge_Chambers_Cost_Model_L_PSH_C96 = (xl_geq(L_per_H__Cost_Model_L_PSH_C95, 7.0) ? "Yes" : "No") # Cost Model L-PSH C96
	@assert xl_compare(Surge_Chambers_Cost_Model_L_PSH_C96, "Yes") # "Cost Model L-PSH!C96"
	# =IF(G15="Yes",IF(O15,O15,IF(C96="Yes",40%,0%)),0)
	LS_Cost_Model_L_PSH_I15 = (xl_eq(Surge_Facilities_Cost_Model_L_PSH_G15, "Yes") ? (xl_logical(NA_Cost_Model_L_PSH_O15) ? NA_Cost_Model_L_PSH_O15 : (xl_eq(Surge_Chambers_Cost_Model_L_PSH_C96, "Yes") ? ((40.0) / 100.0) : ((0.0) / 100.0))) : 0.0) # Cost Model L-PSH I15
	@assert xl_compare(LS_Cost_Model_L_PSH_I15, 0.4) # "Cost Model L-PSH!I15"
	# =I15*SUM(Q20:Q25)
	NA_Cost_Model_L_PSH_Q15 = LS_Cost_Model_L_PSH_I15 * xl_sum(tab_Cost_Model_L_PSH_Q20_R25[!, "Q"]) # Cost Model L-PSH Q15
	@assert xl_compare(NA_Cost_Model_L_PSH_Q15, 1.031014335510007e8) # "Cost Model L-PSH!Q15"
	
	NA_Cost_Model_L_PSH_Q15
end
function calculate_s_Cost_Model_L_PSH_J29(tab_Cost_Model_L_PSH_C97_C105, s_Cost_Curves_L_PSH_D374, s_Market_Adj_Factors_H51, s_Market_Adj_Factors_H58, tab_Market_Adj_Factors_C45_C48)
	# =D374
	s_Cost_Curves_L_PSH_D378 = s_Cost_Curves_L_PSH_D374 # Cost Curves L-PSH D378
	@assert xl_compare(s_Cost_Curves_L_PSH_D378, 1283.3248207034676) # "Cost Curves L-PSH!D378"
	tab_Market_Adj_Factors_C45_C48[3, "C"] = s_Market_Adj_Factors_H58 / s_Market_Adj_Factors_H51 # Market Adj Factors C47 Row: 3
	@assert xl_compare(tab_Market_Adj_Factors_C45_C48[3, "C"], 1.2008673634969647) # "Market Adj Factors!C47"
	# =697.41*D378^-0.329*'Market Adj Factors'!C47
	s_Cost_Curves_L_PSH_E378 = 697.41 * ((s_Cost_Curves_L_PSH_D378) ^ ((-1 * 0.329))) * tab_Market_Adj_Factors_C45_C48[3, "C"] # Cost Curves L-PSH E378
	@assert xl_compare(s_Cost_Curves_L_PSH_E378, 79.49512066200255) # "Cost Curves L-PSH!E378"
	# =1051.9*(D374^-0.204)*'Market Adj Factors'!C47
	s_Cost_Curves_L_PSH_E374 = 1051.9 * ((s_Cost_Curves_L_PSH_D374) ^ ((-1 * 0.204))) * tab_Market_Adj_Factors_C45_C48[3, "C"] # Cost Curves L-PSH E374
	@assert xl_compare(s_Cost_Curves_L_PSH_E374, 293.3383412967246) # "Cost Curves L-PSH!E374"
	# =IF(C104>100,0,IF(C104<=100,'Cost Curves L-PSH'!E374,'Cost Curves L-PSH'!E378))
	s_Cost_Model_L_PSH_J29 = (xl_gt(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 100.0) ? 0.0 : (xl_leq(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 100.0) ? s_Cost_Curves_L_PSH_E374 : s_Cost_Curves_L_PSH_E378)) # Cost Model L-PSH J29
	@assert xl_compare(s_Cost_Model_L_PSH_J29, 0.0) # "Cost Model L-PSH!J29"
	
	s_Cost_Model_L_PSH_J29
end
function calculate_s_Cost_Model_L_PSH_Q37(Switchyard_Cost_Model_L_PSH_G37, s_Cost_Model_L_PSH_P37, tab_Locational_Adj_Factors_A3_B55, Location_Cost_Model_L_PSH_C9, tab_Market_Adj_Factors_C45_C48, Inflation_Factor_Cost_Model_L_PSH_C51, Switchyard_Market_Adj_Factors_C43, Substation__Cost_Model_L_PSH_C63, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, tab_Cost_Model_L_PSH_C97_C105, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10)
	# =-0.1*('Cost Model L-PSH'!C106)^2 + 1.6*('Cost Model L-PSH'!C106) - 1.1
	s_Cost_Curves_L_PSH_M248 = (-1 * 0.1) * ((No_Units__Cost_Model_L_PSH_C106) ^ (2.0)) + 1.6 * No_Units__Cost_Model_L_PSH_C106 - 1.1 # Cost Curves L-PSH M248
	@assert xl_compare(s_Cost_Curves_L_PSH_M248, 3.7000000000000006) # "Cost Curves L-PSH!M248"
	# =-0.0643*('Cost Model L-PSH'!C106)^2 + 1.0743*'Cost Model L-PSH'!C106 - 0.48
	s_Cost_Curves_L_PSH_L248 = (-1 * 0.0643) * ((No_Units__Cost_Model_L_PSH_C106) ^ (2.0)) + 1.0743 * No_Units__Cost_Model_L_PSH_C106 - 0.48 # Cost Curves L-PSH L248
	@assert xl_compare(s_Cost_Curves_L_PSH_L248, 2.7884) # "Cost Curves L-PSH!L248"
	# =-0.6286*('Cost Model L-PSH'!C106)^2 + 8.3086*('Cost Model L-PSH'!C106) - 6.45
	s_Cost_Curves_L_PSH_O248 = (-1 * 0.6286) * ((No_Units__Cost_Model_L_PSH_C106) ^ (2.0)) + 8.3086 * No_Units__Cost_Model_L_PSH_C106 - 6.45 # Cost Curves L-PSH O248
	@assert xl_compare(s_Cost_Curves_L_PSH_O248, 16.7268) # "Cost Curves L-PSH!O248"
	# =-0.1607*('Cost Model L-PSH'!C106)^2 + 3.1007*('Cost Model L-PSH'!C106) - 1.91
	s_Cost_Curves_L_PSH_N248 = (-1 * 0.1607) * ((No_Units__Cost_Model_L_PSH_C106) ^ (2.0)) + 3.1007 * No_Units__Cost_Model_L_PSH_C106 - 1.91 # Cost Curves L-PSH N248
	@assert xl_compare(s_Cost_Curves_L_PSH_N248, 7.921599999999998) # "Cost Curves L-PSH!N248"
	# =IF('Cost Model L-PSH'!C63<=160,'Cost Curves L-PSH'!L248,IF(AND('Cost Model L-PSH'!C63>160,'Cost Model L-PSH'!C63<=230),'Cost Curves L-PSH'!M248,IF(AND('Cost Model L-PSH'!C63>230,'Cost Model L-PSH'!C63<=345),'Cost Curves L-PSH'!N248,IF('Cost Model L-PSH'!C63>345,'Cost Curves L-PSH'!O248))))
	s_Cost_Curves_L_PSH_Q248 = (xl_leq(Substation__Cost_Model_L_PSH_C63, 160.0) ? s_Cost_Curves_L_PSH_L248 : (all([xl_gt(Substation__Cost_Model_L_PSH_C63, 160.0), xl_leq(Substation__Cost_Model_L_PSH_C63, 230.0)]) ? s_Cost_Curves_L_PSH_M248 : (all([xl_gt(Substation__Cost_Model_L_PSH_C63, 230.0), xl_leq(Substation__Cost_Model_L_PSH_C63, 345.0)]) ? s_Cost_Curves_L_PSH_N248 : (xl_gt(Substation__Cost_Model_L_PSH_C63, 345.0) ? s_Cost_Curves_L_PSH_O248 : missing)))) # Cost Curves L-PSH Q248
	@assert xl_compare(s_Cost_Curves_L_PSH_Q248, 16.7268) # "Cost Curves L-PSH!Q248"
	# ='Market Adj Factors'!C43
	s_Cost_Model_L_PSH_M37 = Switchyard_Market_Adj_Factors_C43 # Cost Model L-PSH M37
	@assert xl_compare(s_Cost_Model_L_PSH_M37, 1.0) # "Cost Model L-PSH!M37"
	# ='Cost Curves L-PSH'!Q248*1000000
	s_Cost_Model_L_PSH_J37 = xl_mul(s_Cost_Curves_L_PSH_Q248, 1.0e6) # Cost Model L-PSH J37
	@assert xl_compare(s_Cost_Model_L_PSH_J37, 1.67268e7) # "Cost Model L-PSH!J37"
	# =IF($C$51="Yes",'Market Adj Factors'!$C$45,1)
	s_Cost_Model_L_PSH_L37 = (xl_eq(Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L37
	@assert xl_compare(s_Cost_Model_L_PSH_L37, 2.4063214260550976) # "Cost Model L-PSH!L37"
	# =VLOOKUP($C$9,'Locational Adj Factors'!$A$3:$B$55,2,FALSE)/100
	s_Cost_Model_L_PSH_K37 = xl_div(xl_vlookup(Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K37
	@assert xl_compare(s_Cost_Model_L_PSH_K37, 1.0) # "Cost Model L-PSH!K37"
	# =IF(P37,P37,J37*K37*L37*M37)
	s_Cost_Model_L_PSH_N37 = (xl_logical(s_Cost_Model_L_PSH_P37) ? s_Cost_Model_L_PSH_P37 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J37, s_Cost_Model_L_PSH_K37), s_Cost_Model_L_PSH_L37), s_Cost_Model_L_PSH_M37)) # Cost Model L-PSH N37
	@assert xl_compare(s_Cost_Model_L_PSH_N37, 4.025005722933841e7) # "Cost Model L-PSH!N37"
	# =IF(G37="Yes",1,0)
	LS_Cost_Model_L_PSH_I37 = (xl_eq(Switchyard_Cost_Model_L_PSH_G37, "Yes") ? 1.0 : 0.0) # Cost Model L-PSH I37
	@assert xl_compare(LS_Cost_Model_L_PSH_I37, 1) # "Cost Model L-PSH!I37"
	# =N37*I37
	s_Cost_Model_L_PSH_Q37 = xl_mul(s_Cost_Model_L_PSH_N37, LS_Cost_Model_L_PSH_I37) # Cost Model L-PSH Q37
	@assert xl_compare(s_Cost_Model_L_PSH_Q37, 4.025005722933841e7) # "Cost Model L-PSH!Q37"
	
	s_Cost_Model_L_PSH_Q37
end
function calculate_s_Cost_Model_L_PSH_J28(tab_Cost_Model_L_PSH_C97_C105, tab_Cost_Model_L_PSH_C80_C86, Ac_Ft_to_Cu_Ft_Cost_Model_L_PSH_C26, Generation_Time_Cost_Model_L_PSH_C21, Pump_Time__Cost_Model_L_PSH_C76)
	# =C76*C21
	Pump_Time__Cost_Model_L_PSH_C112 = Pump_Time__Cost_Model_L_PSH_C76 * Generation_Time_Cost_Model_L_PSH_C21 # Cost Model L-PSH C112
	@assert xl_compare(Pump_Time__Cost_Model_L_PSH_C112, 22.2) # "Cost Model L-PSH!C112"
	# =C85*C26/C112/3600
	Mean_Pump_Discharge__Cost_Model_L_PSH_C113 = tab_Cost_Model_L_PSH_C80_C86[6, "C"] * Ac_Ft_to_Cu_Ft_Cost_Model_L_PSH_C26 / Pump_Time__Cost_Model_L_PSH_C112 / 3600.0 # Cost Model L-PSH C113
	@assert xl_compare(Mean_Pump_Discharge__Cost_Model_L_PSH_C113, 8937.273852164664) # "Cost Model L-PSH!C113"
	# ='Cost Model L-PSH'!C113*448.83
	s_Cost_Curves_L_PSH_D363 = Mean_Pump_Discharge__Cost_Model_L_PSH_C113 * 448.83 # Cost Curves L-PSH D363
	@assert xl_compare(s_Cost_Curves_L_PSH_D363, 4.011316623067066e6) # "Cost Curves L-PSH!D363"
	# =0.7799*(D363)^0.7442*1000
	s_Cost_Curves_L_PSH_D365 = 0.7799 * ((s_Cost_Curves_L_PSH_D363) ^ (0.7442)) * 1000.0 # Cost Curves L-PSH D365
	@assert xl_compare(s_Cost_Curves_L_PSH_D365, 6.400369743724546e7) # "Cost Curves L-PSH!D365"
	# =IF(C104>100,0,'Cost Curves L-PSH'!D365)
	s_Cost_Model_L_PSH_J28 = (xl_gt(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 100.0) ? 0.0 : s_Cost_Curves_L_PSH_D365) # Cost Model L-PSH J28
	@assert xl_compare(s_Cost_Model_L_PSH_J28, 0.0) # "Cost Model L-PSH!J28"
	
	s_Cost_Model_L_PSH_J28
end
function calculate_s_Cost_Model_L_PSH_J25(Penstock_Cost_Model_L_PSH_C48, tab_Cost_Model_L_PSH_C97_C105, Nominal_Max_Head_Cost_Model_L_PSH_C14, tab_Cost_Model_L_PSH_C108_C111)
	# =0.0226*('Cost Model L-PSH'!C110)^1.9901*1000
	s_Cost_Curves_L_PSH_M131 = 0.0226 * ((tab_Cost_Model_L_PSH_C108_C111[3, "C"]) ^ (1.9901)) * 1000.0 # Cost Curves L-PSH M131
	@assert xl_compare(s_Cost_Curves_L_PSH_M131, 2911.4757778134895) # "Cost Curves L-PSH!M131"
	# =0.02*('Cost Model L-PSH'!C110)^1.885*1000
	s_Cost_Curves_L_PSH_L131 = 0.02 * ((tab_Cost_Model_L_PSH_C108_C111[3, "C"]) ^ (1.885)) * 1000.0 # Cost Curves L-PSH L131
	@assert xl_compare(s_Cost_Curves_L_PSH_L131, 1993.436465983931) # "Cost Curves L-PSH!L131"
	# =0.0483*('Cost Model L-PSH'!C110)^1.9694*1000
	s_Cost_Curves_L_PSH_O131 = 0.0483 * ((tab_Cost_Model_L_PSH_C108_C111[3, "C"]) ^ (1.9694)) * 1000.0 # Cost Curves L-PSH O131
	@assert xl_compare(s_Cost_Curves_L_PSH_O131, 5915.680161082789) # "Cost Curves L-PSH!O131"
	# =0.0721*('Cost Model L-PSH'!C110)^1.9647*1000
	s_Cost_Curves_L_PSH_P131 = 0.0721 * ((tab_Cost_Model_L_PSH_C108_C111[3, "C"]) ^ (1.9647)) * 1000.0 # Cost Curves L-PSH P131
	@assert xl_compare(s_Cost_Curves_L_PSH_P131, 8729.907486102633) # "Cost Curves L-PSH!P131"
	# =0.0365*('Cost Model L-PSH'!C110)^1.9417*1000
	s_Cost_Curves_L_PSH_N131 = 0.0365 * ((tab_Cost_Model_L_PSH_C108_C111[3, "C"]) ^ (1.9417)) * 1000.0 # Cost Curves L-PSH N131
	@assert xl_compare(s_Cost_Curves_L_PSH_N131, 4178.1254626161035) # "Cost Curves L-PSH!N131"
	# =IF(OR(C104<100, C48="Surface"),IF(C14<=300,'Cost Curves L-PSH'!L131,IF(AND(C14>300,C14<=500),'Cost Curves L-PSH'!M131,IF(AND(C14>500,C14<=700),'Cost Curves L-PSH'!N131,IF(AND(C14>700,C14<=1000),'Cost Curves L-PSH'!O131,IF(C14>1000,'Cost Curves L-PSH'!P131))))),0)
	s_Cost_Model_L_PSH_J25 = (any([xl_lt(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 100.0), xl_eq(Penstock_Cost_Model_L_PSH_C48, "Surface")]) ? (xl_leq(Nominal_Max_Head_Cost_Model_L_PSH_C14, 300.0) ? s_Cost_Curves_L_PSH_L131 : (all([xl_gt(Nominal_Max_Head_Cost_Model_L_PSH_C14, 300.0), xl_leq(Nominal_Max_Head_Cost_Model_L_PSH_C14, 500.0)]) ? s_Cost_Curves_L_PSH_M131 : (all([xl_gt(Nominal_Max_Head_Cost_Model_L_PSH_C14, 500.0), xl_leq(Nominal_Max_Head_Cost_Model_L_PSH_C14, 700.0)]) ? s_Cost_Curves_L_PSH_N131 : (all([xl_gt(Nominal_Max_Head_Cost_Model_L_PSH_C14, 700.0), xl_leq(Nominal_Max_Head_Cost_Model_L_PSH_C14, 1000.0)]) ? s_Cost_Curves_L_PSH_O131 : (xl_gt(Nominal_Max_Head_Cost_Model_L_PSH_C14, 1000.0) ? s_Cost_Curves_L_PSH_P131 : missing))))) : 0.0) # Cost Model L-PSH J25
	@assert xl_compare(s_Cost_Model_L_PSH_J25, 0.0) # "Cost Model L-PSH!J25"
	
	s_Cost_Model_L_PSH_J25
end
function calculate_s_Cost_Model_L_PSH_N39(s_Cost_Model_L_PSH_P39, tab_Market_Adj_Factors_C45_C48, Inflation_Factor_Cost_Model_L_PSH_C51, Transmission_works, tab_Locational_Adj_Factors_A3_B55, Location_Cost_Model_L_PSH_C9, Transmission__Cost_Model_L_PSH_C62, Transmission_Type_num_circuits__Cost_Model_L_PSH_C65, tab_Cost_Curves_L_PSH_H276_I284, Transmission_Terrain___Cost_Model_L_PSH_C64, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, tab_Cost_Model_L_PSH_C97_C105, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10, s_Cost_Curves_L_PSH_J292, s_Cost_Curves_L_PSH_I292, s_Cost_Curves_L_PSH_I291, s_Cost_Curves_L_PSH_J291, s_Cost_Curves_L_PSH_I290, s_Cost_Curves_L_PSH_J290, tab_Cost_Curves_L_PSH_K290_K292)
	tab_Cost_Curves_L_PSH_K290_K292[1, "K"] = s_Cost_Curves_L_PSH_J290 / s_Cost_Curves_L_PSH_I290 # Cost Curves L-PSH K290 Row: 1
	@assert xl_compare(tab_Cost_Curves_L_PSH_K290_K292[1, "K"], 1.6009169532145462) # "Cost Curves L-PSH!K290"
	tab_Cost_Curves_L_PSH_K290_K292[2, "K"] = s_Cost_Curves_L_PSH_J291 / s_Cost_Curves_L_PSH_I291 # Cost Curves L-PSH K291 Row: 2
	@assert xl_compare(tab_Cost_Curves_L_PSH_K290_K292[2, "K"], 1.6001637148385177) # "Cost Curves L-PSH!K291"
	tab_Cost_Curves_L_PSH_K290_K292[3, "K"] = s_Cost_Curves_L_PSH_J292 / s_Cost_Curves_L_PSH_I292 # Cost Curves L-PSH K292 Row: 3
	@assert xl_compare(tab_Cost_Curves_L_PSH_K290_K292[3, "K"], 1.6003282190210737) # "Cost Curves L-PSH!K292"
	# =AVERAGE(K290:K292)
	Double_Cost_Curves_L_PSH_I287 = xl_average(tab_Cost_Curves_L_PSH_K290_K292[!, "K"]) # Cost Curves L-PSH I287
	@assert xl_compare(Double_Cost_Curves_L_PSH_I287, 1.6004696290247125) # "Cost Curves L-PSH!I287"
	# =0.0003*('Cost Model L-PSH'!C105)^2 - 0.2536*('Cost Model L-PSH'!C105) + 335.57
	s_Cost_Curves_L_PSH_D303 = 0.0003 * ((tab_Cost_Model_L_PSH_C97_C105[9, "C"]) ^ (2.0)) - 0.2536 * tab_Cost_Model_L_PSH_C97_C105[9, "C"] + 335.57 # Cost Curves L-PSH D303
	@assert xl_compare(s_Cost_Curves_L_PSH_D303, 504.1956040996767) # "Cost Curves L-PSH!D303"
	# =0.0007*('Cost Model L-PSH'!C105)^2 - 0.0382*('Cost Model L-PSH'!C105) + 169.98
	s_Cost_Curves_L_PSH_C303 = 0.0007 * ((tab_Cost_Model_L_PSH_C97_C105[9, "C"]) ^ (2.0)) - 0.0382 * tab_Cost_Model_L_PSH_C97_C105[9, "C"] + 169.98 # Cost Curves L-PSH C303
	@assert xl_compare(s_Cost_Curves_L_PSH_C303, 1273.8028086526385) # "Cost Curves L-PSH!C303"
	# =0.0043*('Cost Model L-PSH'!C105)^2 - 0.028*'Cost Model L-PSH'!C105 + 108.72
	s_Cost_Curves_L_PSH_B303 = 0.0043 * ((tab_Cost_Model_L_PSH_C97_C105[9, "C"]) ^ (2.0)) - 0.028 * tab_Cost_Model_L_PSH_C97_C105[9, "C"] + 108.72 # Cost Curves L-PSH B303
	@assert xl_compare(s_Cost_Curves_L_PSH_B303, 7154.554065384728) # "Cost Curves L-PSH!B303"
	# =0.00004*('Cost Model L-PSH'!C105)^2 - 0.0652*('Cost Model L-PSH'!C105)+ 428.5
	s_Cost_Curves_L_PSH_E303 = 4.0e-5 * ((tab_Cost_Model_L_PSH_C97_C105[9, "C"]) ^ (2.0)) - 0.0652 * tab_Cost_Model_L_PSH_C97_C105[9, "C"] + 428.5 # Cost Curves L-PSH E303
	@assert xl_compare(s_Cost_Curves_L_PSH_E303, 410.70412550747744) # "Cost Curves L-PSH!E303"
	# =VLOOKUP(C64,'Cost Curves L-PSH'!H276:I284,2,FALSE)
	Tranmission_Terrain_Multiplier___Cost_Model_L_PSH_C117 = xl_vlookup(Transmission_Terrain___Cost_Model_L_PSH_C64, tab_Cost_Curves_L_PSH_H276_I284[!, Between("H", "I")], 2.0, false) # Cost Model L-PSH C117
	@assert xl_compare(Tranmission_Terrain_Multiplier___Cost_Model_L_PSH_C117, 1.75) # "Cost Model L-PSH!C117"
	# =IF(C65="Double",'Cost Curves L-PSH'!I287,1)
	Tranmission_Type_Multiplier__Cost_Model_L_PSH_C118 = (xl_eq(Transmission_Type_num_circuits__Cost_Model_L_PSH_C65, "Double") ? Double_Cost_Curves_L_PSH_I287 : 1.0) # Cost Model L-PSH C118
	@assert xl_compare(Tranmission_Type_Multiplier__Cost_Model_L_PSH_C118, 1.6004696290247125) # "Cost Model L-PSH!C118"
	# =IF('Cost Model L-PSH'!C62<=138,'Cost Curves L-PSH'!B303,IF(AND('Cost Model L-PSH'!C62>138,'Cost Model L-PSH'!C62<=230),'Cost Curves L-PSH'!C303,IF(AND('Cost Model L-PSH'!C62>230,'Cost Model L-PSH'!C62<=345),'Cost Curves L-PSH'!D303,IF('Cost Model L-PSH'!C62<=138>=345,'Cost Curves L-PSH'!E303))))
	s_Cost_Curves_L_PSH_G303 = (xl_leq(Transmission__Cost_Model_L_PSH_C62, 138.0) ? s_Cost_Curves_L_PSH_B303 : (all([xl_gt(Transmission__Cost_Model_L_PSH_C62, 138.0), xl_leq(Transmission__Cost_Model_L_PSH_C62, 230.0)]) ? s_Cost_Curves_L_PSH_C303 : (all([xl_gt(Transmission__Cost_Model_L_PSH_C62, 230.0), xl_leq(Transmission__Cost_Model_L_PSH_C62, 345.0)]) ? s_Cost_Curves_L_PSH_D303 : (xl_leq(Transmission__Cost_Model_L_PSH_C62, xl_geq(138.0, 345.0)) ? s_Cost_Curves_L_PSH_E303 : missing)))) # Cost Curves L-PSH G303
	@assert xl_compare(s_Cost_Curves_L_PSH_G303, 410.70412550747744) # "Cost Curves L-PSH!G303"
	# ='Cost Curves L-PSH'!G303*1000*C117*C118
	s_Cost_Model_L_PSH_J39 = xl_mul(xl_mul(xl_mul(s_Cost_Curves_L_PSH_G303, 1000.0), Tranmission_Terrain_Multiplier___Cost_Model_L_PSH_C117), Tranmission_Type_Multiplier__Cost_Model_L_PSH_C118) # Cost Model L-PSH J39
	@assert xl_compare(s_Cost_Model_L_PSH_J39, 1.1503090889322748e6) # "Cost Model L-PSH!J39"
	# =VLOOKUP($C$9,'Locational Adj Factors'!$A$3:$B$55,2,FALSE)/100
	s_Cost_Model_L_PSH_K39 = xl_div(xl_vlookup(Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K39
	@assert xl_compare(s_Cost_Model_L_PSH_K39, 1.0) # "Cost Model L-PSH!K39"
	# ='Market Adj Factors'!C42
	s_Cost_Model_L_PSH_M39 = Transmission_works # Cost Model L-PSH M39
	@assert xl_compare(s_Cost_Model_L_PSH_M39, 1.3) # "Cost Model L-PSH!M39"
	# =IF($C$51="Yes",'Market Adj Factors'!$C$45,1)
	s_Cost_Model_L_PSH_L39 = (xl_eq(Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L39
	@assert xl_compare(s_Cost_Model_L_PSH_L39, 2.4063214260550976) # "Cost Model L-PSH!L39"
	# =IF(P39,P39,J39*K39*L39*M39)
	s_Cost_Model_L_PSH_N39 = (xl_logical(s_Cost_Model_L_PSH_P39) ? s_Cost_Model_L_PSH_P39 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J39, s_Cost_Model_L_PSH_K39), s_Cost_Model_L_PSH_L39), s_Cost_Model_L_PSH_M39)) # Cost Model L-PSH N39
	@assert xl_compare(s_Cost_Model_L_PSH_N39, 3.598417429468747e6) # "Cost Model L-PSH!N39"
	
	s_Cost_Model_L_PSH_N39
end
function calculate_s_Cost_Model_L_PSH_J30(tab_Cost_Model_L_PSH_C97_C105, Power_Station_Cost_Model_L_PSH_C47, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10, Mean_Gross_Head__Cost_Model_L_PSH_C89)
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0617*('Cost Model L-PSH'!C89) + 241.53
	s_Cost_Curves_L_PSH_D171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0617 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 241.53 # Cost Curves L-PSH D171
	@assert xl_compare(s_Cost_Curves_L_PSH_D171, 194.88132000000002) # "Cost Curves L-PSH!D171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0355*('Cost Model L-PSH'!C89) + 132.8
	s_Cost_Curves_L_PSH_AG171 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0355 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 132.8 # Cost Curves L-PSH AG171
	@assert xl_compare(s_Cost_Curves_L_PSH_AG171, 103.30976000000001) # "Cost Curves L-PSH!AG171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0391*('Cost Model L-PSH'!C89) + 133.75
	s_Cost_Curves_L_PSH_AH213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0391 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 133.75 # Cost Curves L-PSH AH213
	@assert xl_compare(s_Cost_Curves_L_PSH_AH213, 99.48616) # "Cost Curves L-PSH!AH213"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0447*('Cost Model L-PSH'!C89) + 144.85
	s_Cost_Curves_L_PSH_AF213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0447 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 144.85 # Cost Curves L-PSH AF213
	@assert xl_compare(s_Cost_Curves_L_PSH_AF213, 103.16056) # "Cost Curves L-PSH!AF213"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0456*('Cost Model L-PSH'!C89) + 196.43
	s_Cost_Curves_L_PSH_O171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0456 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 196.43 # Cost Curves L-PSH O171
	@assert xl_compare(s_Cost_Curves_L_PSH_O171, 171.12992) # "Cost Curves L-PSH!O171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0313*('Cost Model L-PSH'!C89) + 155.12
	s_Cost_Curves_L_PSH_X213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0313 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 155.12 # Cost Curves L-PSH X213
	@assert xl_compare(s_Cost_Curves_L_PSH_X213, 131.19896) # "Cost Curves L-PSH!X213"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0453*('Cost Model L-PSH'!C89) + 200.29
	s_Cost_Curves_L_PSH_N171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0453 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 200.29 # Cost Curves L-PSH N171
	@assert xl_compare(s_Cost_Curves_L_PSH_N171, 175.38772) # "Cost Curves L-PSH!N171"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0452*('Cost Model L-PSH'!C89) + 157.75
	s_Cost_Curves_L_PSH_Y171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0452 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 157.75 # Cost Curves L-PSH Y171
	@assert xl_compare(s_Cost_Curves_L_PSH_Y171, 132.98032) # "Cost Curves L-PSH!Y171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0448*('Cost Model L-PSH'!C89) + 166.28
	s_Cost_Curves_L_PSH_W171 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0448 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 166.28 # Cost Curves L-PSH W171
	@assert xl_compare(s_Cost_Curves_L_PSH_W171, 124.45796) # "Cost Curves L-PSH!W171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0309*('Cost Model L-PSH'!C89) + 151.13
	s_Cost_Curves_L_PSH_Y213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0309 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 151.13 # Cost Curves L-PSH Y213
	@assert xl_compare(s_Cost_Curves_L_PSH_Y213, 127.73936) # "Cost Curves L-PSH!Y213"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0338*('Cost Model L-PSH'!C89)+ 159.73
	s_Cost_Curves_L_PSH_W213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0338 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 159.73 # Cost Curves L-PSH W213
	@assert xl_compare(s_Cost_Curves_L_PSH_W213, 132.49396) # "Cost Curves L-PSH!W213"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0555*('Cost Model L-PSH'!C89) + 232.3
	s_Cost_Curves_L_PSH_E171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0555 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 232.3 # Cost Curves L-PSH E171
	@assert xl_compare(s_Cost_Curves_L_PSH_E171, 193.87252) # "Cost Curves L-PSH!E171"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0501*('Cost Model L-PSH'!C89)+ 195
	s_Cost_Curves_L_PSH_O213 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0501 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 195.0 # Cost Curves L-PSH O213
	@assert xl_compare(s_Cost_Curves_L_PSH_O213, 163.73292) # "Cost Curves L-PSH!O213"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0609*('Cost Model L-PSH'!C89) + 247.79
	s_Cost_Curves_L_PSH_C171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0609 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 247.79 # Cost Curves L-PSH C171
	@assert xl_compare(s_Cost_Curves_L_PSH_C171, 202.20211999999998) # "Cost Curves L-PSH!C171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.041*('Cost Model L-PSH'!C89)+ 136.46
	s_Cost_Curves_L_PSH_AG213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.041 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 136.46 # Cost Curves L-PSH AG213
	@assert xl_compare(s_Cost_Curves_L_PSH_AG213, 99.67676) # "Cost Curves L-PSH!AG213"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.033*('Cost Model L-PSH'!C89) + 129.39
	s_Cost_Curves_L_PSH_AH171 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.033 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 129.39 # Cost Curves L-PSH AH171
	@assert xl_compare(s_Cost_Curves_L_PSH_AH171, 103.21475999999998) # "Cost Curves L-PSH!AH171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0319*('Cost Model L-PSH'!C89) + 165.52
	s_Cost_Curves_L_PSH_V213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0319 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 165.52 # Cost Curves L-PSH V213
	@assert xl_compare(s_Cost_Curves_L_PSH_V213, 140.80336) # "Cost Curves L-PSH!V213"
	# =0.000008*('Cost Model L-PSH'!C89)^2 - 0.0174*('Cost Model L-PSH'!C89) + 242.74
	s_Cost_Curves_L_PSH_B213 = 8.0e-6 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0174 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 242.74 # Cost Curves L-PSH B213
	@assert xl_compare(s_Cost_Curves_L_PSH_B213, 233.733808) # "Cost Curves L-PSH!B213"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0544*('Cost Model L-PSH'!C89) + 177.56
	s_Cost_Curves_L_PSH_V171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0544 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 177.56 # Cost Curves L-PSH V171
	@assert xl_compare(s_Cost_Curves_L_PSH_V171, 140.59112) # "Cost Curves L-PSH!V171"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.046*('Cost Model L-PSH'!C89) + 196.81
	s_Cost_Curves_L_PSH_N213 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.046 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 196.81 # Cost Curves L-PSH N213
	@assert xl_compare(s_Cost_Curves_L_PSH_N213, 170.97952) # "Cost Curves L-PSH!N213"
	# =0.000008*('Cost Model L-PSH'!C89)^2 - 0.0168*('Cost Model L-PSH'!C89) + 223.7
	s_Cost_Curves_L_PSH_D213 = 8.0e-6 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0168 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 223.7 # Cost Curves L-PSH D213
	@assert xl_compare(s_Cost_Curves_L_PSH_D213, 215.489408) # "Cost Curves L-PSH!D213"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0365*('Cost Model L-PSH'!C89) + 139.89
	s_Cost_Curves_L_PSH_AF171 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0365 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 139.89 # Cost Curves L-PSH AF171
	@assert xl_compare(s_Cost_Curves_L_PSH_AF171, 109.07376) # "Cost Curves L-PSH!AF171"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0457*('Cost Model L-PSH'!C89) + 200.9
	s_Cost_Curves_L_PSH_M213 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0457 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 200.9 # Cost Curves L-PSH M213
	@assert xl_compare(s_Cost_Curves_L_PSH_M213, 175.46732) # "Cost Curves L-PSH!M213"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.048*('Cost Model L-PSH'!C89) + 163.78
	s_Cost_Curves_L_PSH_X171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.048 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 163.78 # Cost Curves L-PSH X171
	@assert xl_compare(s_Cost_Curves_L_PSH_X171, 135.29752) # "Cost Curves L-PSH!X171"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0485*('Cost Model L-PSH'!C89) + 216.18
	s_Cost_Curves_L_PSH_L171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0485 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 216.18 # Cost Curves L-PSH L171
	@assert xl_compare(s_Cost_Curves_L_PSH_L171, 187.03452) # "Cost Curves L-PSH!L171"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0609*'Cost Model L-PSH'!C89 + 259.97
	s_Cost_Curves_L_PSH_B171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0609 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 259.97 # Cost Curves L-PSH B171
	@assert xl_compare(s_Cost_Curves_L_PSH_B171, 214.38212000000004) # "Cost Curves L-PSH!B171"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0406*('Cost Model L-PSH'!C89) + 131.67
	s_Cost_Curves_L_PSH_AI213 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0406 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 131.67 # Cost Curves L-PSH AI213
	@assert xl_compare(s_Cost_Curves_L_PSH_AI213, 95.41716) # "Cost Curves L-PSH!AI213"
	# =0.000008*('Cost Model L-PSH'!C89)^2 - 0.017*('Cost Model L-PSH'!C89) + 230.3
	s_Cost_Curves_L_PSH_C213 = 8.0e-6 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.017 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 230.3 # Cost Curves L-PSH C213
	@assert xl_compare(s_Cost_Curves_L_PSH_C213, 221.824208) # "Cost Curves L-PSH!C213"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0462*('Cost Model L-PSH'!C89) + 206.27
	s_Cost_Curves_L_PSH_M171 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0462 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 206.27 # Cost Curves L-PSH M171
	@assert xl_compare(s_Cost_Curves_L_PSH_M171, 180.17432000000002) # "Cost Curves L-PSH!M171"
	# =0.000008*('Cost Model L-PSH'!C89)^2 - 0.0173*('Cost Model L-PSH'!C89) + 219.28
	s_Cost_Curves_L_PSH_E213 = 8.0e-6 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0173 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 219.28 # Cost Curves L-PSH E213
	@assert xl_compare(s_Cost_Curves_L_PSH_E213, 210.406408) # "Cost Curves L-PSH!E213"
	# =0.00001*('Cost Model L-PSH'!C89)^2 - 0.0316*('Cost Model L-PSH'!C89) + 125.28
	s_Cost_Curves_L_PSH_AI171 = 1.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0316 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 125.28 # Cost Curves L-PSH AI171
	@assert xl_compare(s_Cost_Curves_L_PSH_AI171, 100.96116) # "Cost Curves L-PSH!AI171"
	# =0.00002*('Cost Model L-PSH'!C89)^2 - 0.0499*('Cost Model L-PSH'!C89) + 212.36
	s_Cost_Curves_L_PSH_L213 = 2.0e-5 * ((Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (2.0)) - 0.0499 * Mean_Gross_Head__Cost_Model_L_PSH_C89 + 212.36 # Cost Curves L-PSH L213
	@assert xl_compare(s_Cost_Curves_L_PSH_L213, 181.35812) # "Cost Curves L-PSH!L213"
	# =IF(C104<=100,0,IF(C47="Underground", IF(C107<=80,IF(C106<=2,'Cost Curves L-PSH'!B171,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!C171,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!D171,IF(C106>4,'Cost Curves L-PSH'!E171)))),IF(AND(C107>80,C107<=125),IF(C106<=2,'Cost Curves L-PSH'!L171,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!M171,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!N171,IF(C106>4,'Cost Curves L-PSH'!O171)))),IF(AND(C107>125,C107<=225),IF(C106<=2,'Cost Curves L-PSH'!V171,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!W171,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!X171,IF(C106>4,'Cost Curves L-PSH'!Y171)))),IF(C107>225,IF(C106<=2,'Cost Curves L-PSH'!AF171,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AG171,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AH171,IF(C106>4,'Cost Curves L-PSH'!AI171)))))))),IF(C47="Surface", IF(C107<=80,IF(C106<=2,'Cost Curves L-PSH'!B213,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!C213,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!D213,IF(C106>4,'Cost Curves L-PSH'!E213)))),IF(AND(C107>80,C107<=125),IF(C106<=2,'Cost Curves L-PSH'!L213,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!M213,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!N213,IF(C106>4,'Cost Curves L-PSH'!O213)))),IF(AND(C107>125,C107<=225),IF(C106<=2,'Cost Curves L-PSH'!V213,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!W213,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!X213,IF(C106>4,'Cost Curves L-PSH'!Y213)))),IF(C107>225,IF(C106<=2,'Cost Curves L-PSH'!AF213,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AG213,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AH213,IF(C106>4,'Cost Curves L-PSH'!AI213)))))))))))
	s_Cost_Model_L_PSH_J30 = (xl_leq(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 100.0) ? 0.0 : (xl_eq(Power_Station_Cost_Model_L_PSH_C47, "Underground") ? (xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 80.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_B171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_C171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_D171 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_E171 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 80.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 125.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_L171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_M171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_N171 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_O171 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 125.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 225.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_V171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_W171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_X171 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_Y171 : missing)))) : (xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 225.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_AF171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AG171 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AH171 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AI171 : missing)))) : missing)))) : (xl_eq(Power_Station_Cost_Model_L_PSH_C47, "Surface") ? (xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 80.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_B213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_C213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_D213 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_E213 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 80.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 125.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_L213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_M213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_N213 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_O213 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 125.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 225.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_V213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_W213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_X213 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_Y213 : missing)))) : (xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 225.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_AF213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AG213 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AH213 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AI213 : missing)))) : missing)))) : missing))) # Cost Model L-PSH J30
	@assert xl_compare(s_Cost_Model_L_PSH_J30, 103.21475999999998) # "Cost Model L-PSH!J30"
	
	s_Cost_Model_L_PSH_J30
end
function calculate_s_Cost_Model_L_PSH_J17(Lower_Dam_Crest_Length_Cost_Model_L_PSH_C19, Avg_Lower_Dam_Height_Cost_Model_L_PSH_C18, tab_Cost_Curves_L_PSH_J95_J96)
	# =(0.00009*('Cost Model L-PSH'!C18)^2 + 0.0039*'Cost Model L-PSH'!C18 + 0.0707)*1000
	LUnit_Volume_CY_per_Ft_Cost_Curves_L_PSH_E96 = (9.0e-5 * ((Avg_Lower_Dam_Height_Cost_Model_L_PSH_C18) ^ (2.0)) + 0.0039 * Avg_Lower_Dam_Height_Cost_Model_L_PSH_C18 + 0.0707) * 1000.0 # Cost Curves L-PSH E96
	@assert xl_compare(LUnit_Volume_CY_per_Ft_Cost_Curves_L_PSH_E96, 628.7) # "Cost Curves L-PSH!E96"
	# =E96*'Cost Model L-PSH'!C19
	LTotal_Volume_CY_Cost_Curves_L_PSH_E97 = LUnit_Volume_CY_per_Ft_Cost_Curves_L_PSH_E96 * Lower_Dam_Crest_Length_Cost_Model_L_PSH_C19 # Cost Curves L-PSH E97
	@assert xl_compare(LTotal_Volume_CY_Cost_Curves_L_PSH_E97, 691570) # "Cost Curves L-PSH!E97"
	# =E97/(10^6)
	LVolume_CY___106_Cost_Curves_L_PSH_E98 = LTotal_Volume_CY_Cost_Curves_L_PSH_E97 / ((10.0) ^ (6.0)) # Cost Curves L-PSH E98
	@assert xl_compare(LVolume_CY___106_Cost_Curves_L_PSH_E98, 0.69157) # "Cost Curves L-PSH!E98"
	tab_Cost_Curves_L_PSH_J95_J96[2, "J"] = 7.75204 * LVolume_CY___106_Cost_Curves_L_PSH_E98 / (LVolume_CY___106_Cost_Curves_L_PSH_E98 - 0.049591) # Cost Curves L-PSH J96 Row: 2
	@assert xl_compare(tab_Cost_Curves_L_PSH_J95_J96[2, "J"], 8.350862415748802) # "Cost Curves L-PSH!J96"
	# ='Cost Curves L-PSH'!J96
	s_Cost_Model_L_PSH_J17 = tab_Cost_Curves_L_PSH_J95_J96[2, "J"] # Cost Model L-PSH J17
	@assert xl_compare(s_Cost_Model_L_PSH_J17, 8.350862415748802) # "Cost Model L-PSH!J17"
	
	s_Cost_Model_L_PSH_J17
end
function calculate_s_Cost_Model_L_PSH_J13(Upper_Dam_Crest_Length_Cost_Model_L_PSH_C17, Avg_Upper_Dam_Height_Cost_Model_L_PSH_C16, tab_Cost_Curves_L_PSH_J95_J96)
	# =(0.00009*('Cost Model L-PSH'!C16)^2 + 0.0039*'Cost Model L-PSH'!C16 + 0.0707)*1000
	UUnit_Volume_CY_per_Ft_Cost_Curves_L_PSH_C96 = (9.0e-5 * ((Avg_Upper_Dam_Height_Cost_Model_L_PSH_C16) ^ (2.0)) + 0.0039 * Avg_Upper_Dam_Height_Cost_Model_L_PSH_C16 + 0.0707) * 1000.0 # Cost Curves L-PSH C96
	@assert xl_compare(UUnit_Volume_CY_per_Ft_Cost_Curves_L_PSH_C96, 1834.7) # "Cost Curves L-PSH!C96"
	# =C96*'Cost Model L-PSH'!C17
	UTotal_Volume_CY_Cost_Curves_L_PSH_C97 = UUnit_Volume_CY_per_Ft_Cost_Curves_L_PSH_C96 * Upper_Dam_Crest_Length_Cost_Model_L_PSH_C17 # Cost Curves L-PSH C97
	@assert xl_compare(UTotal_Volume_CY_Cost_Curves_L_PSH_C97, 2385110) # "Cost Curves L-PSH!C97"
	# =C97/(10^6)
	UVolume_CY___106_Cost_Curves_L_PSH_C98 = UTotal_Volume_CY_Cost_Curves_L_PSH_C97 / ((10.0) ^ (6.0)) # Cost Curves L-PSH C98
	@assert xl_compare(UVolume_CY___106_Cost_Curves_L_PSH_C98, 2.38511) # "Cost Curves L-PSH!C98"
	tab_Cost_Curves_L_PSH_J95_J96[1, "J"] = 7.75204 * UVolume_CY___106_Cost_Curves_L_PSH_C98 / (UVolume_CY___106_Cost_Curves_L_PSH_C98 - 0.049591) # Cost Curves L-PSH J95 Row: 1
	@assert xl_compare(tab_Cost_Curves_L_PSH_J95_J96[1, "J"], 7.916642135816494) # "Cost Curves L-PSH!J95"
	# ='Cost Curves L-PSH'!J95
	s_Cost_Model_L_PSH_J13 = tab_Cost_Curves_L_PSH_J95_J96[1, "J"] # Cost Model L-PSH J13
	@assert xl_compare(s_Cost_Model_L_PSH_J13, 7.916642135816494) # "Cost Model L-PSH!J13"
	
	s_Cost_Model_L_PSH_J13
end
function calculate_s_Cost_Model_L_PSH_Q42(s_Cost_Model_L_PSH_L42, s_Cost_Model_L_PSH_M42, s_Cost_Model_L_PSH_P42, Water_Supply__Cost_Model_L_PSH_C40, Water_Supply_Cost_Model_L_PSH_G42, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, tab_Cost_Model_L_PSH_C97_C105, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10, Water_Supply_Cost___Cost_Model_L_PSH_C41, tab_Locational_Adj_Factors_A3_B55, Location_Cost_Model_L_PSH_C9)
	# =VLOOKUP($C$9,'Locational Adj Factors'!$A$3:$B$55,2,FALSE)/100
	s_Cost_Model_L_PSH_K42 = xl_div(xl_vlookup(Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K42
	@assert xl_compare(s_Cost_Model_L_PSH_K42, 1.0) # "Cost Model L-PSH!K42"
	# =C41
	s_Cost_Model_L_PSH_J42 = Water_Supply_Cost___Cost_Model_L_PSH_C41 # Cost Model L-PSH J42
	@assert xl_compare(s_Cost_Model_L_PSH_J42, 200.0) # "Cost Model L-PSH!J42"
	# =IF(G42="Yes",IF(C40="Yes",C105*1000,0),0)
	kW_Cost_Model_L_PSH_I42 = (xl_eq(Water_Supply_Cost_Model_L_PSH_G42, "Yes") ? (xl_eq(Water_Supply__Cost_Model_L_PSH_C40, "Yes") ? tab_Cost_Model_L_PSH_C97_C105[9, "C"] * 1000.0 : 0.0) : 0.0) # Cost Model L-PSH I42
	@assert xl_compare(kW_Cost_Model_L_PSH_I42, 0) # "Cost Model L-PSH!I42"
	# =IF(P42,P42,J42*K42*L42*M42)
	s_Cost_Model_L_PSH_N42 = (xl_logical(s_Cost_Model_L_PSH_P42) ? s_Cost_Model_L_PSH_P42 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J42, s_Cost_Model_L_PSH_K42), s_Cost_Model_L_PSH_L42), s_Cost_Model_L_PSH_M42)) # Cost Model L-PSH N42
	@assert xl_compare(s_Cost_Model_L_PSH_N42, 200.0) # "Cost Model L-PSH!N42"
	# =N42*I42
	s_Cost_Model_L_PSH_Q42 = xl_mul(s_Cost_Model_L_PSH_N42, kW_Cost_Model_L_PSH_I42) # Cost Model L-PSH Q42
	@assert xl_compare(s_Cost_Model_L_PSH_Q42, 0) # "Cost Model L-PSH!Q42"
	
	s_Cost_Model_L_PSH_Q42
end
struct Outputs
	_dollar_per_kWh_Max_Energy_Capacity_Cost_Model_L_PSH_G57
end
@kwdef mutable struct Inputs
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	s_Cost_Curves_L_PSH_I290::Float64 = 959700.0
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	s_Cost_Curves_L_PSH_I291::Float64 = 1.3438e6
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	s_Cost_Curves_L_PSH_I292::Float64 = 1.91945e6
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	s_Cost_Curves_L_PSH_J290::Float64 = 1.5364e6
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	s_Cost_Curves_L_PSH_J291::Float64 = 2.1503e6
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	s_Cost_Curves_L_PSH_J292::Float64 = 3.07175e6
	# used in 1 statements, [s_Cost_Model_L_PSH_J33]
	s_Cost_Curves_L_PSH_V241::String = "Steep"
	# used in 1 statements, [s_Cost_Model_L_PSH_J33]
	s_Cost_Curves_L_PSH_V242::String = "Mild"
	# used in 1 statements, [s_Cost_Model_L_PSH_J33]
	s_Cost_Curves_L_PSH_V243::String = "Flat"
	# used in 1 statements, [s_Cost_Model_L_PSH_J33]
	Terrain_Cost_Curves_L_PSH_W240::String = "New"
	# used in 1 statements, [s_Cost_Model_L_PSH_J33]
	New_Cost_Curves_L_PSH_X240::String = "Rebuild"
	# used in 20 statements, [s_Cost_Model_L_PSH_K10], [tab_Cost_Model_L_PSH_K13_K14[2, "K"]], [tab_Cost_Model_L_PSH_K28_K30[1, "K"]], [tab_Cost_Model_L_PSH_K16_N17[2, "K"]], [tab_Cost_Model_L_PSH_K28_K30[2, "K"]], [tab_Cost_Model_L_PSH_K16_N17[1, "K"]], [tab_Cost_Model_L_PSH_K33_N34[1, "K"]], [tab_Cost_Model_L_PSH_K20_N25[4, "K"]], [tab_Cost_Model_L_PSH_K13_K14[1, "K"]], [tab_Cost_Model_L_PSH_K28_K30[3, "K"]], [tab_Cost_Model_L_PSH_K20_N25[5, "K"]], [tab_Cost_Model_L_PSH_K20_N25[2, "K"]], [s_Cost_Model_L_PSH_J8], [tab_Cost_Model_L_PSH_K20_N25[6, "K"]], [tab_Cost_Model_L_PSH_K20_N25[1, "K"]], [tab_Cost_Model_L_PSH_K33_N34[2, "K"]], [tab_Cost_Model_L_PSH_K20_N25[3, "K"]], [s_Cost_Model_L_PSH_Q37], [s_Cost_Model_L_PSH_N39], [s_Cost_Model_L_PSH_Q42]
	Location_Cost_Model_L_PSH_C9::String = "United States"
	# used in 1 statements, [tab_Cost_Model_L_PSH_C80_C86[2, "C"]]
	Avg_Max_Upper_Reservoir_Depth_Cost_Model_L_PSH_C10::Float64 = 101.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_C80_C86[2, "C"]]
	Upper_Reservoir_Area_Cost_Model_L_PSH_C11::Float64 = 191.0
	# used in 72 statements, [s_Cost_Curves_L_PSH_AK53], [s_Cost_Curves_L_PSH_AG11], [s_Cost_Curves_L_PSH_Q11], [s_Cost_Curves_L_PSH_AM53], [s_Cost_Curves_L_PSH_AL53], [s_Cost_Curves_L_PSH_S53], [s_Cost_Curves_L_PSH_R53], [s_Cost_Curves_L_PSH_Y11], [s_Cost_Curves_L_PSH_F11], [s_Cost_Curves_L_PSH_G53], [s_Cost_Curves_L_PSH_E11], [s_Cost_Curves_L_PSH_D11], [s_Cost_Curves_L_PSH_AA53], [s_Cost_Curves_L_PSH_N53], [s_Cost_Curves_L_PSH_Z53], [s_Cost_Curves_L_PSH_O11], [s_Cost_Curves_L_PSH_W53], [Est__dollar_per_kW_Cost_Curves_L_PSH_B53], [s_Cost_Curves_L_PSH_AJ11], [s_Cost_Curves_L_PSH_AI11], [Min_Gross_Head__Cost_Model_L_PSH_C88], [s_Cost_Curves_L_PSH_AL11], [s_Cost_Curves_L_PSH_E53], [s_Cost_Model_L_PSH_J22], [s_Cost_Curves_L_PSH_AB11], [s_Cost_Curves_L_PSH_I53], [s_Cost_Curves_L_PSH_P53], [s_Cost_Curves_L_PSH_C11], [s_Cost_Curves_L_PSH_Q53], [s_Cost_Curves_L_PSH_M11], [s_Cost_Curves_L_PSH_O53], [s_Cost_Curves_L_PSH_R11], [s_Cost_Curves_L_PSH_W11], [s_Cost_Curves_L_PSH_AH11], [Est__dollar_per_kW_Cost_Curves_L_PSH_L11], [s_Cost_Curves_L_PSH_AA130], [s_Cost_Curves_L_PSH_AI53], [s_Cost_Curves_L_PSH_D53], [s_Cost_Curves_L_PSH_S11], [s_Cost_Curves_L_PSH_AB53], [s_Cost_Curves_L_PSH_P11], [s_Cost_Curves_L_PSH_AG53], [s_Cost_Curves_L_PSH_Y53], [s_Cost_Curves_L_PSH_H11], [Est__dollar_per_kW_Cost_Curves_L_PSH_L53], [s_Cost_Curves_L_PSH_I11], [s_Cost_Curves_L_PSH_AJ53], [s_Cost_Curves_L_PSH_AA11], [Est__dollar_per_kW_Cost_Curves_L_PSH_V53], [s_Cost_Curves_L_PSH_AM11], [s_Cost_Curves_L_PSH_M53], [s_Cost_Curves_L_PSH_X53], [s_Cost_Curves_L_PSH_F53], [Est__dollar_per_kW_Cost_Curves_L_PSH_V11], [s_Cost_Curves_L_PSH_H53], [s_Cost_Curves_L_PSH_N11], [s_Cost_Model_L_PSH_J23], [Est__dollar_per_kW_Cost_Curves_L_PSH_AF11], [s_Cost_Curves_L_PSH_AC11], [Est__dollar_per_kW_Cost_Curves_L_PSH_B11], [Mean_Gross_Head__Cost_Model_L_PSH_C89], [s_Cost_Curves_L_PSH_X11], [Est__dollar_per_kW_Cost_Curves_L_PSH_AF53], [s_Cost_Curves_L_PSH_C53], [s_Cost_Curves_L_PSH_AK11], [s_Cost_Curves_L_PSH_AC53], [s_Cost_Curves_L_PSH_V130], [s_Cost_Curves_L_PSH_AH53], [s_Cost_Curves_L_PSH_G11], [s_Cost_Curves_L_PSH_Z11], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...], [s_Cost_Model_L_PSH_J25]
	Nominal_Max_Head_Cost_Model_L_PSH_C14::Float64 = 1560.0
	# used in 5 statements, [Ft_Cost_Model_L_PSH_I24], [tab_Cost_Model_L_PSH_C97_C105[2, "C"]], [Ft_Cost_Model_L_PSH_I20], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...], [NA_Cost_Model_L_PSH_Q15]
	Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15::Float64 = 14394.0
	# used in 2 statements, [tab_Cost_Model_L_PSH_C80_C86[4, "C"]], [s_Cost_Model_L_PSH_J13]
	Avg_Upper_Dam_Height_Cost_Model_L_PSH_C16::Float64 = 120.0
	# used in 2 statements, [tab_Cost_Model_L_PSH_C80_C86[4, "C"]], [s_Cost_Model_L_PSH_J13]
	Upper_Dam_Crest_Length_Cost_Model_L_PSH_C17::Float64 = 1300.0
	# used in 2 statements, [tab_Cost_Model_L_PSH_C80_C86[5, "C"]], [s_Cost_Model_L_PSH_J17]
	Avg_Lower_Dam_Height_Cost_Model_L_PSH_C18::Float64 = 60.0
	# used in 2 statements, [tab_Cost_Model_L_PSH_C80_C86[5, "C"]], [s_Cost_Model_L_PSH_J17]
	Lower_Dam_Crest_Length_Cost_Model_L_PSH_C19::Float64 = 1100.0
	# used in 1 statements, [Acres_Cost_Model_L_PSH_I8]
	Acreage_to_be_acquired_Cost_Model_L_PSH_C20::Float64 = 2185.0
	# used in 3 statements, [_dollar_per_kWh_Max_Energy_Capacity_Cost_Model_L_PSH_G57], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...], [s_Cost_Model_L_PSH_J28]
	Generation_Time_Cost_Model_L_PSH_C21::Float64 = 18.5
	# used in 2 statements, [tab_Cost_Model_L_PSH_C97_C105[8, "C"]], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Meter_to_Feet_Cost_Model_L_PSH_C24::Float64 = 3.28084
	# used in 2 statements, [s_Cost_Model_L_PSH_J20], [s_Cost_Model_L_PSH_J24]
	Feet_to_Miles_Cost_Model_L_PSH_C25::Float64 = 0.000189394
	# used in 2 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...], [s_Cost_Model_L_PSH_J28]
	Ac_Ft_to_Cu_Ft_Cost_Model_L_PSH_C26::Float64 = 43559.9
	# used in 1 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Hr_to_Sec_Cost_Model_L_PSH_C28::Float64 = 3600.0
	# used in 2 statements, [tab_Cost_Model_L_PSH_C97_C105[8, "C"]], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	MW_to_W_Cost_Model_L_PSH_C29::Float64 = 1.0e6
	# used in 2 statements, [tab_Cost_Model_L_PSH_C97_C105[8, "C"]], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	acceleration_of_gravity_metric_Cost_Model_L_PSH_C30::Float64 = 9.81
	# used in 2 statements, [tab_Cost_Model_L_PSH_C97_C105[8, "C"]], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	density_of_water_metric_Cost_Model_L_PSH_C31::Float64 = 1000.0
	# used in 2 statements, [s_Cost_Model_L_PSH_J20], [s_Cost_Model_L_PSH_J24]
	Tunneling_Condition_Cost_Model_L_PSH_C34::String = "Average"
	# used in 1 statements, [s_Cost_Model_L_PSH_J33]
	Access_Road_Terrain_Cost_Model_L_PSH_C35::String = "Flat"
	# used in 1 statements, [s_Cost_Model_L_PSH_J33]
	Access_Road_Type_Cost_Model_L_PSH_C36::String = "New"
	# used in 1 statements, [pcnt_Cost_Model_L_PSH_I35]
	Highway_Realignment_Cost_Model_L_PSH_C37::String = "Yes"
	# used in 1 statements, [Miles_Cost_Model_L_PSH_I33]
	Access_Road_Cost_Model_L_PSH_C38::Float64 = 3.61
	# used in 2 statements, [Ft_Cost_Model_L_PSH_I34], [s_Cost_Curves_L_PSH_D242]
	Access_Tunnel_Length_Cost_Model_L_PSH_C39::Float64 = 1.25
	# used in 1 statements, [s_Cost_Model_L_PSH_Q42]
	Water_Supply__Cost_Model_L_PSH_C40::String = "No"
	# used in 1 statements, [s_Cost_Model_L_PSH_Q42]
	Water_Supply_Cost___Cost_Model_L_PSH_C41::Float64 = 200.0
	# used in 1 statements, [s_Cost_Model_L_PSH_J14]
	U_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_C44::String = "Vertical"
	# used in 1 statements, [s_Cost_Model_L_PSH_J16]
	L_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_C45::String = "Horizontal"
	# used in 1 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Power_Station_Structure_Geology_Cost_Model_L_PSH_C46::String = "Adverse"
	# used in 7 statements, [Ft_Cost_Model_L_PSH_I34], [s_Cost_Model_L_PSH_J34], [Ft_Cost_Model_L_PSH_I23], [num_Cost_Model_L_PSH_I16], [s_Cost_Model_L_PSH_J23], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...], [s_Cost_Model_L_PSH_J30]
	Power_Station_Cost_Model_L_PSH_C47::String = "Underground"
	# used in 3 statements, [Ft_Cost_Model_L_PSH_I22], [Ft_Cost_Model_L_PSH_I25], [s_Cost_Model_L_PSH_J25]
	Penstock_Cost_Model_L_PSH_C48::String = "Underground"
	# used in 16 statements, [tab_Cost_Model_L_PSH_K33_N34[1, "L"]], [tab_Cost_Model_L_PSH_K20_N25[6, "L"]], [s_Cost_Model_L_PSH_L14], [tab_Cost_Model_L_PSH_K20_N25[2, "L"]], [tab_Cost_Model_L_PSH_K20_N25[1, "L"]], [tab_Cost_Model_L_PSH_K20_N25[4, "L"]], [s_Cost_Model_L_PSH_L10], [tab_Cost_Model_L_PSH_K33_N34[2, "L"]], [tab_Cost_Model_L_PSH_K16_N17[2, "L"]], [tab_Cost_Model_L_PSH_K20_N25[3, "L"]], [s_Cost_Model_L_PSH_L30], [s_Cost_Model_L_PSH_L8], [tab_Cost_Model_L_PSH_K20_N25[5, "L"]], [tab_Cost_Model_L_PSH_K16_N17[1, "L"]], [s_Cost_Model_L_PSH_Q37], [s_Cost_Model_L_PSH_N39]
	Inflation_Factor_Cost_Model_L_PSH_C51::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[1, "I"]]
	Mobilization_per_Demobilization__Cost_Model_L_PSH_C52::Float64 = 0.05
	# used in 3 statements, [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q50], [NA_Cost_Model_L_PSH_Q46]
	Material_per_Equipment_pcnt__Cost_Model_L_PSH_C53::Float64 = 1.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[2, "I"]]
	Sales_Tax__Cost_Model_L_PSH_C54::Float64 = 0.06
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[3, "I"]]
	Contingency__Cost_Model_L_PSH_C55::Float64 = 0.33
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[4, "I"]]
	EPC_Cost__Cost_Model_L_PSH_C56::Float64 = 0.25
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[5, "I"]]
	Developer_Cost__Cost_Model_L_PSH_C57::Float64 = 0.03
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[6, "I"]]
	Overhead__and__Profit__Cost_Model_L_PSH_C58::Float64 = 0.07
	# used in 1 statements, [Miles_Cost_Model_L_PSH_I39]
	Transmission_Distance__Cost_Model_L_PSH_C61::Float64 = 13.5
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	Transmission__Cost_Model_L_PSH_C62::Float64 = 500.0
	# used in 1 statements, [s_Cost_Model_L_PSH_Q37]
	Substation__Cost_Model_L_PSH_C63::Float64 = 500.0
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	Transmission_Terrain___Cost_Model_L_PSH_C64::String = "Mountain"
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	Transmission_Type_num_circuits__Cost_Model_L_PSH_C65::String = "Double"
	# used in 1 statements, [tab_Cost_Model_L_PSH_C80_C86[6, "C"]]
	Active_Storage__Cost_Model_L_PSH_C68::Float64 = 0.85
	# used in 1 statements, [Min_Gross_Head__Cost_Model_L_PSH_C88]
	Hmin_per_Hmax__Cost_Model_L_PSH_C69::Float64 = 0.7
	# used in 2 statements, [tab_Cost_Model_L_PSH_C97_C105[2, "C"]], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Hazen_Williams_C__Cost_Model_L_PSH_C70::Float64 = 90.0
	# used in 2 statements, [tab_Cost_Model_L_PSH_C97_C105[8, "C"]], [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	P_T_Efficiency__Cost_Model_L_PSH_C71::Float64 = 0.88
	# used in 1 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Max_Tunnel_Velocity__Cost_Model_L_PSH_C72::Float64 = 23.0
	# used in 1 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Max_Tunnel_Dia__Cost_Model_L_PSH_C73::Float64 = 35.0
	# used in 1 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Max_Unit_Capacity__Cost_Model_L_PSH_C74::Float64 = 350.0
	# used in 1 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	Min_No_Units__Cost_Model_L_PSH_C75::Float64 = 2.0
	# used in 1 statements, [s_Cost_Model_L_PSH_J28]
	Pump_Time__Cost_Model_L_PSH_C76::Float64 = 12.0 / 10.0
	# used in 1 statements, [Acres_Cost_Model_L_PSH_I8]
	Land_and_Land_Rights_Cost_Model_L_PSH_G8::String = "Yes"
	# used in 1 statements, [kW_Cost_Model_L_PSH_I10]
	Powerplant_Structure_Cost_Model_L_PSH_G10::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I13_I14[1, "I"]]
	Upper_Reservoir_Dam_and_Spillway_Cost_Model_L_PSH_G13::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I13_I14[2, "I"]]
	Upper_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_G14::String = "Yes"
	# used in 1 statements, [NA_Cost_Model_L_PSH_Q15]
	Surge_Facilities_Cost_Model_L_PSH_G15::String = "Yes"
	# used in 1 statements, [num_Cost_Model_L_PSH_I16]
	Lower_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_G16::String = "Yes"
	# used in 1 statements, [CY_Cost_Model_L_PSH_I17]
	Lower_Reservoir_Dam_and_Spillway_Cost_Model_L_PSH_G17::String = "Yes"
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I20]
	Upper_Low__and__High_Pressure_Tunnels_Cost_Model_L_PSH_G20::String = "Yes"
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I21]
	Vertical_Shafts_Cost_Model_L_PSH_G21::String = "Yes"
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I22]
	Penstock_Tunnels_Cost_Model_L_PSH_G22::String = "Yes"
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I23]
	Draft_Tube_Tunnels_Cost_Model_L_PSH_G23::String = "Yes"
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I24]
	Tailrace_Tunnels_Cost_Model_L_PSH_G24::String = "Yes"
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I25]
	Surface_Penstock_Cost_Model_L_PSH_G25::String = "Yes"
	# used in 1 statements, [LS_Cost_Model_L_PSH_I28]
	Pump_per_Motors_Cost_Model_L_PSH_G28::String = "Yes"
	# used in 1 statements, [kW_Cost_Model_L_PSH_I29]
	Generator_per_Turbines_Cost_Model_L_PSH_G29::String = "Yes"
	# used in 1 statements, [kW_Cost_Model_L_PSH_I30]
	Total_Powerstation_Cost_Model_L_PSH_G30::String = "Yes"
	# used in 1 statements, [Miles_Cost_Model_L_PSH_I33]
	Access_Roads_Cost_Model_L_PSH_G33::String = "Yes"
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I34]
	Access_Tunnels_Cost_Model_L_PSH_G34::String = "Yes"
	# used in 1 statements, [pcnt_Cost_Model_L_PSH_I35]
	Highway_Realignment_Cost_Model_L_PSH_G35::String = "Yes"
	# used in 1 statements, [s_Cost_Model_L_PSH_Q37]
	Switchyard_Cost_Model_L_PSH_G37::String = "Yes"
	# used in 1 statements, [Miles_Cost_Model_L_PSH_I39]
	Transmission_Lines_Cost_Model_L_PSH_G39::String = "Yes"
	# used in 1 statements, [s_Cost_Model_L_PSH_Q42]
	Water_Supply_Cost_Model_L_PSH_G42::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[1, "I"]]
	Mobilization_per_Demobilization_Cost_Model_L_PSH_G45::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[2, "I"]]
	Sales_Tax_Cost_Model_L_PSH_G46::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[3, "I"]]
	Contingency_Cost_Model_L_PSH_G47::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[4, "I"]]
	EPC_Cost_Cost_Model_L_PSH_G48::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[5, "I"]]
	Developer_Cost_Cost_Model_L_PSH_G49::String = "Yes"
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[6, "I"]]
	Overhead__and__Profit_Cost_Model_L_PSH_G50::String = "Yes"
	# used in 1 statements, [s_Cost_Model_L_PSH_N8]
	s_Cost_Model_L_PSH_K8::Float64 = 1.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N28_N30[1, "N"]]
	s_Cost_Model_L_PSH_L28::Float64 = 1.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N28_N30[2, "N"]]
	s_Cost_Model_L_PSH_L29::Float64 = 1.0
	# used in 1 statements, [s_Cost_Model_L_PSH_Q42]
	s_Cost_Model_L_PSH_L42::Float64 = 1.0
	# used in 1 statements, [s_Cost_Model_L_PSH_N8]
	s_Cost_Model_L_PSH_M8::Float64 = 1.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N28_N30[1, "N"]]
	s_Cost_Model_L_PSH_M28::Float64 = 1.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N28_N30[2, "N"]]
	s_Cost_Model_L_PSH_M29::Float64 = 1.0
	# used in 1 statements, [s_Cost_Model_L_PSH_Q42]
	s_Cost_Model_L_PSH_M42::Float64 = 1.0
	# used in 1 statements, [Acres_Cost_Model_L_PSH_I8]
	s_Cost_Model_L_PSH_O8::Float64 = 0.0
	# used in 1 statements, [kW_Cost_Model_L_PSH_I10]
	s_Cost_Model_L_PSH_O10::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I13_I14[1, "I"]]
	s_Cost_Model_L_PSH_O13::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I13_I14[2, "I"]]
	s_Cost_Model_L_PSH_O14::Float64 = 0.0
	# used in 1 statements, [NA_Cost_Model_L_PSH_Q15]
	NA_Cost_Model_L_PSH_O15::Float64 = 0.0
	# used in 1 statements, [num_Cost_Model_L_PSH_I16]
	s_Cost_Model_L_PSH_O16::Float64 = 0.0
	# used in 1 statements, [CY_Cost_Model_L_PSH_I17]
	s_Cost_Model_L_PSH_O17::Float64 = 0.0
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I20]
	s_Cost_Model_L_PSH_O20::Missing = missing
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I21]
	s_Cost_Model_L_PSH_O21::Missing = missing
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I22]
	s_Cost_Model_L_PSH_O22::Missing = missing
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I23]
	s_Cost_Model_L_PSH_O23::Missing = missing
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I24]
	s_Cost_Model_L_PSH_O24::Missing = missing
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I25]
	s_Cost_Model_L_PSH_O25::Float64 = 0.0
	# used in 1 statements, [kW_Cost_Model_L_PSH_I29]
	s_Cost_Model_L_PSH_O29::Float64 = 0.0
	# used in 1 statements, [kW_Cost_Model_L_PSH_I30]
	s_Cost_Model_L_PSH_O30::Float64 = 0.0
	# used in 1 statements, [Miles_Cost_Model_L_PSH_I33]
	s_Cost_Model_L_PSH_O33::Float64 = 0.0
	# used in 1 statements, [Ft_Cost_Model_L_PSH_I34]
	s_Cost_Model_L_PSH_O34::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[1, "I"]]
	s_Cost_Model_L_PSH_O45::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[2, "I"]]
	s_Cost_Model_L_PSH_O46::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[3, "I"]]
	s_Cost_Model_L_PSH_O47::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[4, "I"]]
	s_Cost_Model_L_PSH_O48::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[5, "I"]]
	s_Cost_Model_L_PSH_O49::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_I45_I50[6, "I"]]
	s_Cost_Model_L_PSH_O50::Float64 = 0.0
	# used in 1 statements, [s_Cost_Model_L_PSH_N8]
	s_Cost_Model_L_PSH_P8::Float64 = 0.0
	# used in 1 statements, [Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91[2, "C"...]
	s_Cost_Model_L_PSH_P10::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N13_N14[1, "N"]]
	s_Cost_Model_L_PSH_P13::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N13_N14[2, "N"]]
	s_Cost_Model_L_PSH_P14::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K16_N17[1, "N"]]
	s_Cost_Model_L_PSH_P16::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K16_N17[2, "N"]]
	s_Cost_Model_L_PSH_P17::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[1, "N"]]
	s_Cost_Model_L_PSH_P20::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[2, "N"]]
	s_Cost_Model_L_PSH_P21::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[3, "N"]]
	s_Cost_Model_L_PSH_P22::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[4, "N"]]
	s_Cost_Model_L_PSH_P23::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[5, "N"]]
	s_Cost_Model_L_PSH_P24::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[6, "N"]]
	s_Cost_Model_L_PSH_P25::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N28_N30[1, "N"]]
	s_Cost_Model_L_PSH_P28::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N28_N30[2, "N"]]
	s_Cost_Model_L_PSH_P29::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_N28_N30[3, "N"]]
	s_Cost_Model_L_PSH_P30::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K33_N34[1, "N"]]
	s_Cost_Model_L_PSH_P33::Float64 = 0.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K33_N34[2, "N"]]
	s_Cost_Model_L_PSH_P34::Float64 = 0.0
	# used in 1 statements, [s_Cost_Model_L_PSH_Q37]
	s_Cost_Model_L_PSH_P37::Float64 = 0.0
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	s_Cost_Model_L_PSH_P39::Float64 = 0.0
	# used in 1 statements, [s_Cost_Model_L_PSH_Q42]
	s_Cost_Model_L_PSH_P42::Float64 = 0.0
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q9::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q11::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q12::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q18::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q19::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q26::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q27::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q31::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q32::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q36::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q38::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q40::Missing = missing
	# used in 7 statements, [NA_Cost_Model_L_PSH_Q45], [NA_Cost_Model_L_PSH_Q49], [NA_Cost_Model_L_PSH_Q48], [NA_Cost_Model_L_PSH_Q47], [NA_Cost_Model_L_PSH_Q50], [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]], [NA_Cost_Model_L_PSH_Q46]
	s_Cost_Model_L_PSH_Q41::Missing = missing
	# used in 1 statements, [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]]
	s_Cost_Model_L_PSH_Q43::Missing = missing
	# used in 1 statements, [tab_Cost_Model_L_PSH_Q52_S52[1, "Q"]]
	s_Cost_Model_L_PSH_Q44::Missing = missing
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A3::String = "Connecticut"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A4::String = "Delaware"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A5::String = "Maine"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A6::String = "Maryland"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A7::String = "Massachusetts"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A8::String = "New Hampshire"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A9::String = "New Jersey"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A10::String = "New York"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A11::String = "Pennsylvania"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A12::String = "Rhode Island"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A13::String = "Vermont"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A14::String = "Michigan"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A15::String = "Minnesota"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A16::String = "Wisconsin"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A17::String = "Illinois"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A18::String = "Indiana"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A19::String = "Iowa"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A20::String = "Missouri"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A21::String = "Ohio"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A22::String = "Kansas"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A23::String = "Nebraska"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A24::String = "North Dakota"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A25::String = "South Dakota"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A26::String = "Kentucky"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A27::String = "North Carolina"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A28::String = "Tennessee"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A29::String = "Virginia"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A30::String = "West Virginia"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A31::String = "Alabama"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A32::String = "Florida"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A33::String = "Georgia"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A34::String = "South Carolina"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A35::String = "Arkansas"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A36::String = "Louisiana"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A37::String = "Mississippi"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A38::String = "Oklahoma"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A39::String = "Texas"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A40::String = "Arizona"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A41::String = "Colorado"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A42::String = "Idaho"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A43::String = "Montana"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A44::String = "Nevada"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A45::String = "New Mexico"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A46::String = "Utah"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A47::String = "Wyoming"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A48::String = "California"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A49::String = "Oregon"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A50::String = "Washington"
	# used in 1 statements, [s_Cost_Model_L_PSH_J8]
	s_Land_Value_Data_A51::String = "United States"
	# used in 1 statements, [s_Cost_Model_L_PSH_M10]
	Power_station::Float64 = 1.3
	# used in 2 statements, [tab_Cost_Model_L_PSH_L13_M13[1, "M"]], [tab_Cost_Model_L_PSH_K16_N17[2, "M"]]
	Dams_spillways_diversions_emb::Float64 = 1.4
	# used in 1 statements, [s_Cost_Model_L_PSH_M14]
	Upper_intake_per_outlet_vertical::Float64 = 2.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K16_N17[1, "M"]]
	Lower_intake_per_outlet_horizontal::Float64 = 1.3
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[6, "M"]]
	Surface_penstocks::Float64 = 1.3
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[2, "M"]]
	Vertical_shaft::Float64 = 1.8
	# used in 2 statements, [tab_Cost_Model_L_PSH_K20_N25[1, "M"]], [tab_Cost_Model_L_PSH_K20_N25[5, "M"]]
	Concrete_lined_tunnels_eg_tailrace__and__power_tun::Float64 = 1.6
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[3, "M"]]
	Steel_Lined_tunnels_underground_penstocks::Float64 = 1.9
	# used in 1 statements, [tab_Cost_Model_L_PSH_K20_N25[4, "M"]]
	Draft_tubes::Float64 = 1.9
	# used in 1 statements, [tab_Cost_Model_L_PSH_K33_N34[2, "M"]]
	Access_and_voltage_tunnels::Float64 = 1.0
	# used in 1 statements, [tab_Cost_Model_L_PSH_K33_N34[1, "M"]]
	Roads::Float64 = 1.0
	# used in 1 statements, [s_Cost_Model_L_PSH_M30]
	Electro_Mechanical_::Float64 = 1.7
	# used in 1 statements, [s_Cost_Model_L_PSH_N39]
	Transmission_works::Float64 = 1.3
	# used in 1 statements, [s_Cost_Model_L_PSH_Q37]
	Switchyard_Market_Adj_Factors_C43::Float64 = 1.0
	# used in 1 statements, [tab_Market_Adj_Factors_C45_C48[1, "C"]]
	s_Market_Adj_Factors_H24::Float64 = 118.27500000000002
	# used in 1 statements, [s_Cost_Model_L_PSH_J29]
	s_Market_Adj_Factors_H51::Float64 = 237.00175
	# used in 2 statements, [tab_Market_Adj_Factors_C45_C48[1, "C"]], [s_Cost_Model_L_PSH_J29]
	s_Market_Adj_Factors_H58::Float64 = 284.6076666666667
end
struct Tables
	tab_Cost_Curves_L_PSH_F87_H93::DataFrame
	tab_Cost_Curves_L_PSH_J95_J96::DataFrame
	tab_Cost_Curves_L_PSH_B131_E132::DataFrame
	tab_Cost_Curves_L_PSH_K290_K292::DataFrame
	tab_Cost_Model_S_PSH_I13_I14::DataFrame
	tab_Cost_Model_S_PSH_K13_K14::DataFrame
	tab_Cost_Model_S_PSH_L13_M13::DataFrame
	tab_Cost_Model_S_PSH_N13_N14::DataFrame
	tab_Cost_Model_S_PSH_Q13_Q14::DataFrame
	tab_Cost_Model_S_PSH_R13_R17::DataFrame
	tab_Cost_Model_S_PSH_K16_N17::DataFrame
	tab_Cost_Model_S_PSH_Q16_Q17::DataFrame
	tab_Cost_Model_S_PSH_K20_N25::DataFrame
	tab_Cost_Model_S_PSH_Q20_R25::DataFrame
	tab_Cost_Model_S_PSH_K28_K30::DataFrame
	tab_Cost_Model_S_PSH_N28_N30::DataFrame
	tab_Cost_Model_S_PSH_Q28_R30::DataFrame
	tab_Cost_Model_S_PSH_K33_N34::DataFrame
	tab_Cost_Model_S_PSH_Q33_R35::DataFrame
	tab_Cost_Model_S_PSH_I45_I50::DataFrame
	tab_Cost_Model_S_PSH_R45_R50::DataFrame
	tab_Cost_Model_S_PSH_Q52_S52::DataFrame
	tab_Cost_Model_S_PSH_C80_C86::DataFrame
	tab_Cost_Model_S_PSH_C90_C91::DataFrame
	tab_Cost_Model_S_PSH_C97_C105::DataFrame
	tab_Cost_Model_S_PSH_C108_C111::DataFrame
	tab_Sensitivity_Analysis_C73_E85::DataFrame
	tab_Sensitivity_Analysis_H73_H85::DataFrame
	tab_Sensitivity_Analysis_C92_E99::DataFrame
	tab_Sensitivity_Analysis_H92_H99::DataFrame
	tab_Land_Value_Data_G3_G51::DataFrame
	tab_Cost_Model_L_PSH_I13_I14::DataFrame
	tab_Cost_Model_L_PSH_K13_K14::DataFrame
	tab_Cost_Model_L_PSH_L13_M13::DataFrame
	tab_Cost_Model_L_PSH_N13_N14::DataFrame
	tab_Cost_Model_L_PSH_Q13_Q14::DataFrame
	tab_Cost_Model_L_PSH_R13_R17::DataFrame
	tab_Cost_Model_L_PSH_K16_N17::DataFrame
	tab_Cost_Model_L_PSH_Q16_Q17::DataFrame
	tab_Cost_Model_L_PSH_K20_N25::DataFrame
	tab_Cost_Model_L_PSH_Q20_R25::DataFrame
	tab_Cost_Model_L_PSH_K28_K30::DataFrame
	tab_Cost_Model_L_PSH_N28_N30::DataFrame
	tab_Cost_Model_L_PSH_Q28_R30::DataFrame
	tab_Cost_Model_L_PSH_K33_N34::DataFrame
	tab_Cost_Model_L_PSH_Q33_R35::DataFrame
	tab_Cost_Model_L_PSH_I45_I50::DataFrame
	tab_Cost_Model_L_PSH_R45_R50::DataFrame
	tab_Cost_Model_L_PSH_Q52_S52::DataFrame
	tab_Cost_Model_L_PSH_C80_C86::DataFrame
	tab_Cost_Model_L_PSH_C90_C91::DataFrame
	tab_Cost_Model_L_PSH_C97_C105::DataFrame
	tab_Cost_Model_L_PSH_C108_C111::DataFrame
	tab_Cost_Curves_S_PSH_F87_H93::DataFrame
	tab_Cost_Curves_S_PSH_J95_J96::DataFrame
	tab_Cost_Curves_S_PSH_K290_K292::DataFrame
	tab_Market_Adj_Factors_I6_I58::DataFrame
	tab_Market_Adj_Factors_D30_D43::DataFrame
	tab_Market_Adj_Factors_C45_C48::DataFrame
	tab_Land_Value_Data_B44_F44::DataFrame
	tab_Land_Value_Data_B8_F8::DataFrame
	tab_Land_Value_Data_B17_F17::DataFrame
	tab_Locational_Adj_Factors_A3_B55::DataFrame
	tab_Land_Value_Data_B38_F38::DataFrame
	tab_Land_Value_Data_B12_F12::DataFrame
	tab_Land_Value_Data_B20_F20::DataFrame
	tab_Land_Value_Data_B36_F36::DataFrame
	tab_Land_Value_Data_B4_F4::DataFrame
	tab_Land_Value_Data_B43_F43::DataFrame
	tab_Land_Value_Data_B11_F11::DataFrame
	tab_Land_Value_Data_B25_F25::DataFrame
	tab_Land_Value_Data_B5_F5::DataFrame
	tab_Land_Value_Data_B31_F31::DataFrame
	tab_Land_Value_Data_B46_F46::DataFrame
	tab_Land_Value_Data_B34_F34::DataFrame
	tab_Land_Value_Data_B32_F32::DataFrame
	tab_Land_Value_Data_B18_F18::DataFrame
	tab_Land_Value_Data_B42_F42::DataFrame
	tab_Land_Value_Data_B49_F49::DataFrame
	tab_Land_Value_Data_B15_F15::DataFrame
	tab_Land_Value_Data_B14_F14::DataFrame
	tab_Land_Value_Data_B35_F35::DataFrame
	tab_Land_Value_Data_B7_F7::DataFrame
	tab_Land_Value_Data_B41_F41::DataFrame
	tab_Land_Value_Data_B33_F33::DataFrame
	tab_Land_Value_Data_B45_F45::DataFrame
	tab_Land_Value_Data_B22_F22::DataFrame
	tab_Land_Value_Data_B47_F47::DataFrame
	tab_Land_Value_Data_B50_F50::DataFrame
	tab_Land_Value_Data_B9_F9::DataFrame
	tab_Land_Value_Data_B30_F30::DataFrame
	tab_Land_Value_Data_B27_F27::DataFrame
	tab_Land_Value_Data_B16_F16::DataFrame
	tab_Land_Value_Data_B28_F28::DataFrame
	tab_Land_Value_Data_B21_F21::DataFrame
	tab_Land_Value_Data_B48_F48::DataFrame
	tab_Land_Value_Data_B37_F37::DataFrame
	tab_Land_Value_Data_B39_F39::DataFrame
	tab_Land_Value_Data_B6_F6::DataFrame
	tab_Land_Value_Data_B51_F51::DataFrame
	tab_Land_Value_Data_B23_F23::DataFrame
	tab_Land_Value_Data_B10_F10::DataFrame
	tab_Cost_Curves_L_PSH_W241_X243::DataFrame
	tab_Land_Value_Data_B29_F29::DataFrame
	tab_Land_Value_Data_B13_F13::DataFrame
	tab_Land_Value_Data_B26_F26::DataFrame
	tab_Land_Value_Data_B3_F3::DataFrame
	tab_Land_Value_Data_B24_F24::DataFrame
	tab_Land_Value_Data_B40_F40::DataFrame
	tab_Land_Value_Data_B19_F19::DataFrame
	tab_Cost_Curves_L_PSH_H276_I284::DataFrame
end

function make_input_tables()
	tab_Cost_Curves_L_PSH_F87_H93 = DataFrame("F" => zeros(7), "G" => zeros(7), "H" => Vector{Any}(missing, 7))
	tab_Cost_Curves_L_PSH_J95_J96 = DataFrame("J" => zeros(2))
	tab_Cost_Curves_L_PSH_B131_E132 = DataFrame("B" => zeros(2), "C" => zeros(2), "D" => zeros(2), "E" => zeros(2))
	tab_Cost_Curves_L_PSH_K290_K292 = DataFrame("K" => zeros(3))
	tab_Cost_Model_S_PSH_I13_I14 = DataFrame("I" => zeros(2))
	tab_Cost_Model_S_PSH_K13_K14 = DataFrame("K" => Vector{Any}(missing, 2))
	tab_Cost_Model_S_PSH_L13_M13 = DataFrame("L" => zeros(1), "M" => zeros(1))
	tab_Cost_Model_S_PSH_N13_N14 = DataFrame("N" => Vector{Any}(missing, 2))
	tab_Cost_Model_S_PSH_Q13_Q14 = DataFrame("Q" => Vector{Any}(missing, 2))
	tab_Cost_Model_S_PSH_R13_R17 = DataFrame("R" => Vector{Any}(missing, 5))
	tab_Cost_Model_S_PSH_K16_N17 = DataFrame("K" => Vector{Any}(missing, 2), "L" => zeros(2), "M" => zeros(2), "N" => Vector{Any}(missing, 2))
	tab_Cost_Model_S_PSH_Q16_Q17 = DataFrame("Q" => Vector{Any}(missing, 2))
	tab_Cost_Model_S_PSH_K20_N25 = DataFrame("K" => Vector{Any}(missing, 6), "L" => zeros(6), "M" => zeros(6), "N" => Vector{Any}(missing, 6))
	tab_Cost_Model_S_PSH_Q20_R25 = DataFrame("Q" => Vector{Any}(missing, 6), "R" => Vector{Any}(missing, 6))
	tab_Cost_Model_S_PSH_K28_K30 = DataFrame("K" => Vector{Any}(missing, 3))
	tab_Cost_Model_S_PSH_N28_N30 = DataFrame("N" => Vector{Any}(missing, 3))
	tab_Cost_Model_S_PSH_Q28_R30 = DataFrame("Q" => Vector{Any}(missing, 3), "R" => Vector{Any}(missing, 3))
	tab_Cost_Model_S_PSH_K33_N34 = DataFrame("K" => Vector{Any}(missing, 2), "L" => zeros(2), "M" => zeros(2), "N" => Vector{Any}(missing, 2))
	tab_Cost_Model_S_PSH_Q33_R35 = DataFrame("Q" => Vector{Any}(missing, 3), "R" => Vector{Any}(missing, 3))
	tab_Cost_Model_S_PSH_I45_I50 = DataFrame("I" => zeros(6))
	tab_Cost_Model_S_PSH_R45_R50 = DataFrame("R" => zeros(6))
	tab_Cost_Model_S_PSH_Q52_S52 = DataFrame("Q" => zeros(1), "R" => zeros(1), "S" => zeros(1))
	tab_Cost_Model_S_PSH_C80_C86 = DataFrame("C" => zeros(7))
	tab_Cost_Model_S_PSH_C90_C91 = DataFrame("C" => zeros(2))
	tab_Cost_Model_S_PSH_C97_C105 = DataFrame("C" => zeros(9))
	tab_Cost_Model_S_PSH_C108_C111 = DataFrame("C" => zeros(4))
	tab_Sensitivity_Analysis_C73_E85 = DataFrame("C" => zeros(13), "D" => zeros(13), "E" => zeros(13))
	tab_Sensitivity_Analysis_H73_H85 = DataFrame("H" => zeros(13))
	tab_Sensitivity_Analysis_C92_E99 = DataFrame("C" => zeros(8), "D" => zeros(8), "E" => zeros(8))
	tab_Sensitivity_Analysis_H92_H99 = DataFrame("H" => zeros(8))
	tab_Land_Value_Data_G3_G51 = DataFrame("G" => zeros(49))
	tab_Cost_Model_L_PSH_I13_I14 = DataFrame("I" => zeros(2))
	tab_Cost_Model_L_PSH_K13_K14 = DataFrame("K" => Vector{Any}(missing, 2))
	tab_Cost_Model_L_PSH_L13_M13 = DataFrame("L" => zeros(1), "M" => zeros(1))
	tab_Cost_Model_L_PSH_N13_N14 = DataFrame("N" => Vector{Any}(missing, 2))
	tab_Cost_Model_L_PSH_Q13_Q14 = DataFrame("Q" => Vector{Any}(missing, 2))
	tab_Cost_Model_L_PSH_R13_R17 = DataFrame("R" => Vector{Any}(missing, 5))
	tab_Cost_Model_L_PSH_K16_N17 = DataFrame("K" => Vector{Any}(missing, 2), "L" => zeros(2), "M" => zeros(2), "N" => Vector{Any}(missing, 2))
	tab_Cost_Model_L_PSH_Q16_Q17 = DataFrame("Q" => Vector{Any}(missing, 2))
	tab_Cost_Model_L_PSH_K20_N25 = DataFrame("K" => Vector{Any}(missing, 6), "L" => zeros(6), "M" => zeros(6), "N" => Vector{Any}(missing, 6))
	tab_Cost_Model_L_PSH_Q20_R25 = DataFrame("Q" => Vector{Any}(missing, 6), "R" => Vector{Any}(missing, 6))
	tab_Cost_Model_L_PSH_K28_K30 = DataFrame("K" => Vector{Any}(missing, 3))
	tab_Cost_Model_L_PSH_N28_N30 = DataFrame("N" => Vector{Any}(missing, 3))
	tab_Cost_Model_L_PSH_Q28_R30 = DataFrame("Q" => Vector{Any}(missing, 3), "R" => Vector{Any}(missing, 3))
	tab_Cost_Model_L_PSH_K33_N34 = DataFrame("K" => Vector{Any}(missing, 2), "L" => zeros(2), "M" => zeros(2), "N" => Vector{Any}(missing, 2))
	tab_Cost_Model_L_PSH_Q33_R35 = DataFrame("Q" => Vector{Any}(missing, 3), "R" => Vector{Any}(missing, 3))
	tab_Cost_Model_L_PSH_I45_I50 = DataFrame("I" => zeros(6))
	tab_Cost_Model_L_PSH_R45_R50 = DataFrame("R" => zeros(6))
	tab_Cost_Model_L_PSH_Q52_S52 = DataFrame("Q" => zeros(1), "R" => zeros(1), "S" => zeros(1))
	tab_Cost_Model_L_PSH_C80_C86 = DataFrame("C" => zeros(7))
	tab_Cost_Model_L_PSH_C90_C91 = DataFrame("C" => zeros(2))
	tab_Cost_Model_L_PSH_C97_C105 = DataFrame("C" => zeros(9))
	tab_Cost_Model_L_PSH_C108_C111 = DataFrame("C" => zeros(4))
	tab_Cost_Curves_S_PSH_F87_H93 = DataFrame("F" => zeros(7), "G" => zeros(7), "H" => Vector{Any}(missing, 7))
	tab_Cost_Curves_S_PSH_J95_J96 = DataFrame("J" => zeros(2))
	tab_Cost_Curves_S_PSH_K290_K292 = DataFrame("K" => zeros(3))
	tab_Market_Adj_Factors_I6_I58 = DataFrame("I" => zeros(53))
	tab_Market_Adj_Factors_D30_D43 = DataFrame("D" => zeros(14))
	tab_Market_Adj_Factors_C45_C48 = DataFrame("C" => zeros(4))
	tab_Land_Value_Data_B44_F44 = DataFrame("B" => zeros(1), "C" => Vector{Missing}(missing, 1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => Vector{Missing}(missing, 1))
	tab_Land_Value_Data_B8_F8 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B17_F17 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Locational_Adj_Factors_A3_B55 = DataFrame("A" => Vector{Union{String, Missing}}(missing, 53), "B" => zeros(53))
	tab_Land_Value_Data_B38_F38 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B12_F12 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B20_F20 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B36_F36 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B4_F4 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B43_F43 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B11_F11 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B25_F25 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B5_F5 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B31_F31 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B46_F46 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B34_F34 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B32_F32 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B18_F18 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B42_F42 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B49_F49 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B15_F15 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B14_F14 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B35_F35 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B7_F7 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B41_F41 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B33_F33 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B45_F45 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B22_F22 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B47_F47 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B50_F50 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B9_F9 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B30_F30 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B27_F27 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B16_F16 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B28_F28 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B21_F21 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B48_F48 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B37_F37 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B39_F39 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B6_F6 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => Vector{Missing}(missing, 1))
	tab_Land_Value_Data_B51_F51 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B23_F23 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => zeros(1), "F" => zeros(1))
	tab_Land_Value_Data_B10_F10 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Cost_Curves_L_PSH_W241_X243 = DataFrame("W" => zeros(3), "X" => zeros(3))
	tab_Land_Value_Data_B29_F29 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B13_F13 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B26_F26 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B3_F3 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B24_F24 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Land_Value_Data_B40_F40 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => zeros(1), "E" => Vector{Missing}(missing, 1), "F" => Vector{Missing}(missing, 1))
	tab_Land_Value_Data_B19_F19 = DataFrame("B" => zeros(1), "C" => zeros(1), "D" => Vector{Missing}(missing, 1), "E" => Vector{Missing}(missing, 1), "F" => zeros(1))
	tab_Cost_Curves_L_PSH_H276_I284 = DataFrame("H" => Vector{Union{String, Missing}}(missing, 9), "I" => zeros(9))

	tab_Land_Value_Data_B44_F44[1, "B"] = 1010.0 # Land Value Data B44 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B44_F44[1, "B"], 1010) # "Land Value Data!B44"
	# "Land Value Data!C44":"Land Value Data!F44"
	@. tab_Land_Value_Data_B44_F44[!, Between("C", "F")] = missing
	tab_Land_Value_Data_B8_F8[!, Between("B", "D")] .= [5050.0 8770.0 missing]
	tab_Land_Value_Data_B8_F8[1, "F"] = 6870.0 # Land Value Data F8 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B8_F8[1, "F"], 6870) # "Land Value Data!F8"
	# "Land Value Data!B17":"Land Value Data!C17"
	@. tab_Land_Value_Data_B17_F17[!, Between("B", "C")] = 7900.0
	# "Land Value Data!D17":"Land Value Data!E17"
	@. tab_Land_Value_Data_B17_F17[!, Between("D", "E")] = missing
	tab_Land_Value_Data_B17_F17[1, "F"] = 3400.0 # Land Value Data F17 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B17_F17[1, "F"], 3400) # "Land Value Data!F17"
	tab_Locational_Adj_Factors_A3_B55[!, "A"] .= ["Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "DC", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming", "Puerto Rico", "United States"]
	tab_Locational_Adj_Factors_A3_B55[1:43, "B"] .= [96.8, 125.3, 97.5, 95.3, 98.2, 101.2, 101.3, 99.3, 102.4, 99.8, 97.2, 111.8, 102.5, 96.9, 96.0, 99.1, 98.8, 94.4, 99.1, 96.6, 97.8, 98.8, 95.5, 98.8, 97.7, 98.0, 102.3, 98.5, 98.4, 97.8, 99.7, 98.1, 99.7, 98.2, 102.0, 96.9, 95.6, 97.6, 95.7, 99.7, 97.1, 100.2, 97.5]
	tab_Locational_Adj_Factors_A3_B55[45:53, "B"] .= [99.5, 97.8, 99.8, 101.9, 99.0, 98.9, 98.4, 121.3, 100.0]
	tab_Land_Value_Data_B38_F38[!, Between("B", "F")] .= [2020.0 1810.0 missing 1790.0 1600.0]
	tab_Land_Value_Data_B12_F12[!, Between("B", "D")] .= [16400.0 8770.0 missing]
	tab_Land_Value_Data_B12_F12[1, "F"] = 6870.0 # Land Value Data F12 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B12_F12[1, "F"], 6870) # "Land Value Data!F12"
	tab_Land_Value_Data_B20_F20[!, Between("B", "F")] .= [3700.0 3810.0 4800.0 3700.0 2160.0]
	tab_Land_Value_Data_B36_F36[!, Between("B", "F")] .= [3220.0 2980.0 2880.0 3020.0 2950.0]
	tab_Land_Value_Data_B4_F4[!, Between("B", "D")] .= [9300.0 8600.0 missing]
	tab_Land_Value_Data_B4_F4[1, "F"] = 6870.0 # Land Value Data F4 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B4_F4[1, "F"], 6870) # "Land Value Data!F4"
	tab_Land_Value_Data_B43_F43[!, Between("B", "F")] .= [930.0 1050.0 3050.0 835.0 700.0]
	tab_Land_Value_Data_B11_F11[!, Between("B", "D")] .= [6800.0 7600.0 missing]
	tab_Land_Value_Data_B11_F11[1, "F"] = 3440.0 # Land Value Data F11 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B11_F11[1, "F"], 3440) # "Land Value Data!F11"
	tab_Land_Value_Data_B25_F25[!, Between("B", "F")] .= [2190.0 3390.0 missing 3180.0 1060.0]
	tab_Land_Value_Data_B5_F5[!, Between("B", "D")] .= [2600.0 8770.0 missing]
	tab_Land_Value_Data_B5_F5[1, "F"] = 6870.0 # Land Value Data F5 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B5_F5[1, "F"], 6870) # "Land Value Data!F5"
	tab_Land_Value_Data_B31_F31[!, Between("B", "D")] .= [3200.0 3550.0 missing]
	tab_Land_Value_Data_B31_F31[1, "F"] = 2650.0 # Land Value Data F31 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B31_F31[1, "F"], 2650) # "Land Value Data!F31"
	tab_Land_Value_Data_B46_F46[!, Between("B", "F")] .= [2620.0 4190.0 6650.0 1550.0 1370.0]
	tab_Land_Value_Data_B34_F34[!, Between("B", "D")] .= [3600.0 2900.0 missing]
	tab_Land_Value_Data_B34_F34[1, "F"] = 3350.0 # Land Value Data F34 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B34_F34[1, "F"], 3350) # "Land Value Data!F34"
	tab_Land_Value_Data_B32_F32[!, Between("B", "F")] .= [6020.0 7300.0 8350.0 6320.0 5530.0]
	tab_Land_Value_Data_B18_F18[!, Between("B", "D")] .= [7100.0 6800.0 missing]
	tab_Land_Value_Data_B18_F18[1, "F"] = 2490.0 # Land Value Data F18 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B18_F18[1, "F"], 2490) # "Land Value Data!F18"
	tab_Land_Value_Data_B42_F42[!, Between("B", "F")] .= [3350.0 4450.0 6800.0 1890.0 1700.0]
	tab_Land_Value_Data_B49_F49[!, Between("B", "F")] .= [2790.0 3310.0 5800.0 2340.0 830.0]
	tab_Land_Value_Data_B15_F15[!, Between("B", "D")] .= [5240.0 5270.0 missing]
	tab_Land_Value_Data_B15_F15[1, "F"] = 1830.0 # Land Value Data F15 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B15_F15[1, "F"], 1830) # "Land Value Data!F15"
	tab_Land_Value_Data_B14_F14[!, Between("B", "D")] .= [5300.0 4700.0 missing]
	tab_Land_Value_Data_B14_F14[1, "F"] = 2740.0 # Land Value Data F14 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B14_F14[1, "F"], 2740) # "Land Value Data!F14"
	tab_Land_Value_Data_B35_F35[!, Between("B", "F")] .= [3390.0 2930.0 3420.0 2130.0 2700.0]
	tab_Land_Value_Data_B7_F7[!, Between("B", "D")] .= [13700.0 8770.0 missing]
	tab_Land_Value_Data_B7_F7[1, "F"] = 6870.0 # Land Value Data F7 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B7_F7[1, "F"], 6870) # "Land Value Data!F7"
	tab_Land_Value_Data_B41_F41[!, Between("B", "F")] .= [1610.0 2240.0 5400.0 1400.0 875.0]
	tab_Land_Value_Data_B33_F33[!, Between("B", "F")] .= [3670.0 3480.0 4350.0 3130.0 4060.0]
	tab_Land_Value_Data_B45_F45[!, Between("B", "F")] .= [600.0 1660.0 4550.0 485.0 440.0]
	tab_Land_Value_Data_B22_F22[!, Between("B", "F")] .= [2100.0 2370.0 3700.0 2250.0 1500.0]
	tab_Land_Value_Data_B47_F47[!, Between("B", "F")] .= [790.0 1600.0 2550.0 890.0 610.0]
	tab_Land_Value_Data_B50_F50[!, Between("B", "F")] .= [2900.0 2700.0 7800.0 1310.0 750.0]
	tab_Land_Value_Data_B9_F9[!, Between("B", "D")] .= [14400.0 14800.0 missing]
	tab_Land_Value_Data_B9_F9[1, "F"] = 13400.0 # Land Value Data F9 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B9_F9[1, "F"], 13400) # "Land Value Data!F9"
	tab_Land_Value_Data_B30_F30[!, Between("B", "D")] .= [2770.0 3330.0 missing]
	tab_Land_Value_Data_B30_F30[1, "F"] = 2200.0 # Land Value Data F30 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B30_F30[1, "F"], 2200) # "Land Value Data!F30"
	tab_Land_Value_Data_B27_F27[!, Between("B", "D")] .= [4750.0 4290.0 missing]
	tab_Land_Value_Data_B27_F27[1, "F"] = 4850.0 # Land Value Data F27 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B27_F27[1, "F"], 4850) # "Land Value Data!F27"
	tab_Land_Value_Data_B16_F16[!, Between("B", "D")] .= [5190.0 5280.0 missing]
	tab_Land_Value_Data_B16_F16[1, "F"] = 2520.0 # Land Value Data F16 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B16_F16[1, "F"], 2520) # "Land Value Data!F16"
	tab_Land_Value_Data_B28_F28[!, Between("B", "D")] .= [4260.0 4130.0 missing]
	tab_Land_Value_Data_B28_F28[1, "F"] = 4000.0 # Land Value Data F28 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B28_F28[1, "F"], 4000) # "Land Value Data!F28"
	tab_Land_Value_Data_B21_F21[!, Between("B", "D")] .= [6600.0 6800.0 missing]
	tab_Land_Value_Data_B21_F21[1, "F"] = 3440.0 # Land Value Data F21 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B21_F21[1, "F"], 3440) # "Land Value Data!F21"
	tab_Land_Value_Data_B48_F48[!, Between("B", "F")] .= [10900.0 13860.0 16300.0 5900.0 3100.0]
	tab_Land_Value_Data_B37_F37[!, Between("B", "F")] .= [2860.0 3150.0 3700.0 2830.0 2480.0]
	tab_Land_Value_Data_B39_F39[!, Between("B", "F")] .= [2380.0 2150.0 2540.0 2090.0 1800.0]
	tab_Land_Value_Data_B6_F6[!, Between("B", "D")] .= [8670.0 7960.0 missing]
	tab_Land_Value_Data_B51_F51[!, Between("B", "D")] .= [3380.0 4420.0 missing]
	tab_Land_Value_Data_B51_F51[1, "F"] = 1480.0 # Land Value Data F51 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B51_F51[1, "F"], 1480) # "Land Value Data!F51"
	tab_Land_Value_Data_B23_F23[!, Between("B", "F")] .= [3100.0 4960.0 6530.0 3990.0 1080.0]
	tab_Land_Value_Data_B10_F10[!, Between("B", "D")] .= [3270.0 2910.0 missing]
	tab_Land_Value_Data_B10_F10[1, "F"] = 1580.0 # Land Value Data F10 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B10_F10[1, "F"], 1580) # "Land Value Data!F10"
	tab_Cost_Curves_L_PSH_W241_X243[!, Between("W", "X")] .= [439000.0 293000.0;283000.0 186000.0;189000.0 123000.0]
	tab_Land_Value_Data_B29_F29[!, Between("B", "D")] .= [4700.0 4790.0 missing]
	tab_Land_Value_Data_B29_F29[1, "F"] = 4060.0 # Land Value Data F29 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B29_F29[1, "F"], 4060) # "Land Value Data!F29"
	tab_Land_Value_Data_B13_F13[!, Between("B", "D")] .= [3900.0 8770.0 missing]
	tab_Land_Value_Data_B13_F13[1, "F"] = 6870.0 # Land Value Data F13 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B13_F13[1, "F"], 6870) # "Land Value Data!F13"
	tab_Land_Value_Data_B26_F26[!, Between("B", "D")] .= [4000.0 4510.0 missing]
	tab_Land_Value_Data_B26_F26[1, "F"] = 3080.0 # Land Value Data F26 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B26_F26[1, "F"], 3080) # "Land Value Data!F26"
	tab_Land_Value_Data_B3_F3[!, Between("B", "D")] .= [12500.0 8770.0 missing]
	tab_Land_Value_Data_B3_F3[1, "F"] = 6870.0 # Land Value Data F3 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B3_F3[1, "F"], 6870) # "Land Value Data!F3"
	tab_Land_Value_Data_B24_F24[!, Between("B", "D")] .= [1820.0 2060.0 missing]
	tab_Land_Value_Data_B24_F24[1, "F"] = 840.0 # Land Value Data F24 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B24_F24[1, "F"], 840) # "Land Value Data!F24"
	tab_Land_Value_Data_B40_F40[1, "B"] = 3900.0 # Land Value Data B40 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B40_F40[1, "B"], 3900) # "Land Value Data!B40"
	# "Land Value Data!C40":"Land Value Data!D40"
	@. tab_Land_Value_Data_B40_F40[!, Between("C", "D")] = 7700.0
	# "Land Value Data!E40":"Land Value Data!F40"
	@. tab_Land_Value_Data_B40_F40[!, Between("E", "F")] = missing
	tab_Land_Value_Data_B19_F19[!, Between("B", "D")] .= [7740.0 7810.0 missing]
	tab_Land_Value_Data_B19_F19[1, "F"] = 3020.0 # Land Value Data F19 Row: 1
	@assert xl_compare(tab_Land_Value_Data_B19_F19[1, "F"], 3020) # "Land Value Data!F19"
	tab_Cost_Curves_L_PSH_H276_I284[!, "H"] .= ["Desert", "Flat", "Farmland", "Forested", "Rolling Hill", "Mountain", "Wetland", "Suburban", "Urban"]
	tab_Cost_Curves_L_PSH_H276_I284[1, "I"] = 1.05 # Cost Curves L-PSH I276 Row: 1
	@assert xl_compare(tab_Cost_Curves_L_PSH_H276_I284[1, "I"], 1.05) # "Cost Curves L-PSH!I276"
	# "Cost Curves L-PSH!I277":"Cost Curves L-PSH!I278"
	@. tab_Cost_Curves_L_PSH_H276_I284[2:3, "I"] = 1.0
	tab_Cost_Curves_L_PSH_H276_I284[4:9, "I"] .= [2.25, 1.4, 1.75, 1.2, 1.27, 1.59]

	Tables(
		tab_Cost_Curves_L_PSH_F87_H93,
		tab_Cost_Curves_L_PSH_J95_J96,
		tab_Cost_Curves_L_PSH_B131_E132,
		tab_Cost_Curves_L_PSH_K290_K292,
		tab_Cost_Model_S_PSH_I13_I14,
		tab_Cost_Model_S_PSH_K13_K14,
		tab_Cost_Model_S_PSH_L13_M13,
		tab_Cost_Model_S_PSH_N13_N14,
		tab_Cost_Model_S_PSH_Q13_Q14,
		tab_Cost_Model_S_PSH_R13_R17,
		tab_Cost_Model_S_PSH_K16_N17,
		tab_Cost_Model_S_PSH_Q16_Q17,
		tab_Cost_Model_S_PSH_K20_N25,
		tab_Cost_Model_S_PSH_Q20_R25,
		tab_Cost_Model_S_PSH_K28_K30,
		tab_Cost_Model_S_PSH_N28_N30,
		tab_Cost_Model_S_PSH_Q28_R30,
		tab_Cost_Model_S_PSH_K33_N34,
		tab_Cost_Model_S_PSH_Q33_R35,
		tab_Cost_Model_S_PSH_I45_I50,
		tab_Cost_Model_S_PSH_R45_R50,
		tab_Cost_Model_S_PSH_Q52_S52,
		tab_Cost_Model_S_PSH_C80_C86,
		tab_Cost_Model_S_PSH_C90_C91,
		tab_Cost_Model_S_PSH_C97_C105,
		tab_Cost_Model_S_PSH_C108_C111,
		tab_Sensitivity_Analysis_C73_E85,
		tab_Sensitivity_Analysis_H73_H85,
		tab_Sensitivity_Analysis_C92_E99,
		tab_Sensitivity_Analysis_H92_H99,
		tab_Land_Value_Data_G3_G51,
		tab_Cost_Model_L_PSH_I13_I14,
		tab_Cost_Model_L_PSH_K13_K14,
		tab_Cost_Model_L_PSH_L13_M13,
		tab_Cost_Model_L_PSH_N13_N14,
		tab_Cost_Model_L_PSH_Q13_Q14,
		tab_Cost_Model_L_PSH_R13_R17,
		tab_Cost_Model_L_PSH_K16_N17,
		tab_Cost_Model_L_PSH_Q16_Q17,
		tab_Cost_Model_L_PSH_K20_N25,
		tab_Cost_Model_L_PSH_Q20_R25,
		tab_Cost_Model_L_PSH_K28_K30,
		tab_Cost_Model_L_PSH_N28_N30,
		tab_Cost_Model_L_PSH_Q28_R30,
		tab_Cost_Model_L_PSH_K33_N34,
		tab_Cost_Model_L_PSH_Q33_R35,
		tab_Cost_Model_L_PSH_I45_I50,
		tab_Cost_Model_L_PSH_R45_R50,
		tab_Cost_Model_L_PSH_Q52_S52,
		tab_Cost_Model_L_PSH_C80_C86,
		tab_Cost_Model_L_PSH_C90_C91,
		tab_Cost_Model_L_PSH_C97_C105,
		tab_Cost_Model_L_PSH_C108_C111,
		tab_Cost_Curves_S_PSH_F87_H93,
		tab_Cost_Curves_S_PSH_J95_J96,
		tab_Cost_Curves_S_PSH_K290_K292,
		tab_Market_Adj_Factors_I6_I58,
		tab_Market_Adj_Factors_D30_D43,
		tab_Market_Adj_Factors_C45_C48,
		tab_Land_Value_Data_B44_F44,
		tab_Land_Value_Data_B8_F8,
		tab_Land_Value_Data_B17_F17,
		tab_Locational_Adj_Factors_A3_B55,
		tab_Land_Value_Data_B38_F38,
		tab_Land_Value_Data_B12_F12,
		tab_Land_Value_Data_B20_F20,
		tab_Land_Value_Data_B36_F36,
		tab_Land_Value_Data_B4_F4,
		tab_Land_Value_Data_B43_F43,
		tab_Land_Value_Data_B11_F11,
		tab_Land_Value_Data_B25_F25,
		tab_Land_Value_Data_B5_F5,
		tab_Land_Value_Data_B31_F31,
		tab_Land_Value_Data_B46_F46,
		tab_Land_Value_Data_B34_F34,
		tab_Land_Value_Data_B32_F32,
		tab_Land_Value_Data_B18_F18,
		tab_Land_Value_Data_B42_F42,
		tab_Land_Value_Data_B49_F49,
		tab_Land_Value_Data_B15_F15,
		tab_Land_Value_Data_B14_F14,
		tab_Land_Value_Data_B35_F35,
		tab_Land_Value_Data_B7_F7,
		tab_Land_Value_Data_B41_F41,
		tab_Land_Value_Data_B33_F33,
		tab_Land_Value_Data_B45_F45,
		tab_Land_Value_Data_B22_F22,
		tab_Land_Value_Data_B47_F47,
		tab_Land_Value_Data_B50_F50,
		tab_Land_Value_Data_B9_F9,
		tab_Land_Value_Data_B30_F30,
		tab_Land_Value_Data_B27_F27,
		tab_Land_Value_Data_B16_F16,
		tab_Land_Value_Data_B28_F28,
		tab_Land_Value_Data_B21_F21,
		tab_Land_Value_Data_B48_F48,
		tab_Land_Value_Data_B37_F37,
		tab_Land_Value_Data_B39_F39,
		tab_Land_Value_Data_B6_F6,
		tab_Land_Value_Data_B51_F51,
		tab_Land_Value_Data_B23_F23,
		tab_Land_Value_Data_B10_F10,
		tab_Cost_Curves_L_PSH_W241_X243,
		tab_Land_Value_Data_B29_F29,
		tab_Land_Value_Data_B13_F13,
		tab_Land_Value_Data_B26_F26,
		tab_Land_Value_Data_B3_F3,
		tab_Land_Value_Data_B24_F24,
		tab_Land_Value_Data_B40_F40,
		tab_Land_Value_Data_B19_F19,
		tab_Cost_Curves_L_PSH_H276_I284,
	)
end
function calculate(inputs::Inputs, tables::Tables)
tab_Cost_Curves_L_PSH_F87_H93 = tables.tab_Cost_Curves_L_PSH_F87_H93
tab_Cost_Curves_L_PSH_J95_J96 = tables.tab_Cost_Curves_L_PSH_J95_J96
tab_Cost_Curves_L_PSH_B131_E132 = tables.tab_Cost_Curves_L_PSH_B131_E132
tab_Cost_Curves_L_PSH_K290_K292 = tables.tab_Cost_Curves_L_PSH_K290_K292
tab_Cost_Model_S_PSH_I13_I14 = tables.tab_Cost_Model_S_PSH_I13_I14
tab_Cost_Model_S_PSH_K13_K14 = tables.tab_Cost_Model_S_PSH_K13_K14
tab_Cost_Model_S_PSH_L13_M13 = tables.tab_Cost_Model_S_PSH_L13_M13
tab_Cost_Model_S_PSH_N13_N14 = tables.tab_Cost_Model_S_PSH_N13_N14
tab_Cost_Model_S_PSH_Q13_Q14 = tables.tab_Cost_Model_S_PSH_Q13_Q14
tab_Cost_Model_S_PSH_R13_R17 = tables.tab_Cost_Model_S_PSH_R13_R17
tab_Cost_Model_S_PSH_K16_N17 = tables.tab_Cost_Model_S_PSH_K16_N17
tab_Cost_Model_S_PSH_Q16_Q17 = tables.tab_Cost_Model_S_PSH_Q16_Q17
tab_Cost_Model_S_PSH_K20_N25 = tables.tab_Cost_Model_S_PSH_K20_N25
tab_Cost_Model_S_PSH_Q20_R25 = tables.tab_Cost_Model_S_PSH_Q20_R25
tab_Cost_Model_S_PSH_K28_K30 = tables.tab_Cost_Model_S_PSH_K28_K30
tab_Cost_Model_S_PSH_N28_N30 = tables.tab_Cost_Model_S_PSH_N28_N30
tab_Cost_Model_S_PSH_Q28_R30 = tables.tab_Cost_Model_S_PSH_Q28_R30
tab_Cost_Model_S_PSH_K33_N34 = tables.tab_Cost_Model_S_PSH_K33_N34
tab_Cost_Model_S_PSH_Q33_R35 = tables.tab_Cost_Model_S_PSH_Q33_R35
tab_Cost_Model_S_PSH_I45_I50 = tables.tab_Cost_Model_S_PSH_I45_I50
tab_Cost_Model_S_PSH_R45_R50 = tables.tab_Cost_Model_S_PSH_R45_R50
tab_Cost_Model_S_PSH_Q52_S52 = tables.tab_Cost_Model_S_PSH_Q52_S52
tab_Cost_Model_S_PSH_C80_C86 = tables.tab_Cost_Model_S_PSH_C80_C86
tab_Cost_Model_S_PSH_C90_C91 = tables.tab_Cost_Model_S_PSH_C90_C91
tab_Cost_Model_S_PSH_C97_C105 = tables.tab_Cost_Model_S_PSH_C97_C105
tab_Cost_Model_S_PSH_C108_C111 = tables.tab_Cost_Model_S_PSH_C108_C111
tab_Sensitivity_Analysis_C73_E85 = tables.tab_Sensitivity_Analysis_C73_E85
tab_Sensitivity_Analysis_H73_H85 = tables.tab_Sensitivity_Analysis_H73_H85
tab_Sensitivity_Analysis_C92_E99 = tables.tab_Sensitivity_Analysis_C92_E99
tab_Sensitivity_Analysis_H92_H99 = tables.tab_Sensitivity_Analysis_H92_H99
tab_Land_Value_Data_G3_G51 = tables.tab_Land_Value_Data_G3_G51
tab_Cost_Model_L_PSH_I13_I14 = tables.tab_Cost_Model_L_PSH_I13_I14
tab_Cost_Model_L_PSH_K13_K14 = tables.tab_Cost_Model_L_PSH_K13_K14
tab_Cost_Model_L_PSH_L13_M13 = tables.tab_Cost_Model_L_PSH_L13_M13
tab_Cost_Model_L_PSH_N13_N14 = tables.tab_Cost_Model_L_PSH_N13_N14
tab_Cost_Model_L_PSH_Q13_Q14 = tables.tab_Cost_Model_L_PSH_Q13_Q14
tab_Cost_Model_L_PSH_R13_R17 = tables.tab_Cost_Model_L_PSH_R13_R17
tab_Cost_Model_L_PSH_K16_N17 = tables.tab_Cost_Model_L_PSH_K16_N17
tab_Cost_Model_L_PSH_Q16_Q17 = tables.tab_Cost_Model_L_PSH_Q16_Q17
tab_Cost_Model_L_PSH_K20_N25 = tables.tab_Cost_Model_L_PSH_K20_N25
tab_Cost_Model_L_PSH_Q20_R25 = tables.tab_Cost_Model_L_PSH_Q20_R25
tab_Cost_Model_L_PSH_K28_K30 = tables.tab_Cost_Model_L_PSH_K28_K30
tab_Cost_Model_L_PSH_N28_N30 = tables.tab_Cost_Model_L_PSH_N28_N30
tab_Cost_Model_L_PSH_Q28_R30 = tables.tab_Cost_Model_L_PSH_Q28_R30
tab_Cost_Model_L_PSH_K33_N34 = tables.tab_Cost_Model_L_PSH_K33_N34
tab_Cost_Model_L_PSH_Q33_R35 = tables.tab_Cost_Model_L_PSH_Q33_R35
tab_Cost_Model_L_PSH_I45_I50 = tables.tab_Cost_Model_L_PSH_I45_I50
tab_Cost_Model_L_PSH_R45_R50 = tables.tab_Cost_Model_L_PSH_R45_R50
tab_Cost_Model_L_PSH_Q52_S52 = tables.tab_Cost_Model_L_PSH_Q52_S52
tab_Cost_Model_L_PSH_C80_C86 = tables.tab_Cost_Model_L_PSH_C80_C86
tab_Cost_Model_L_PSH_C90_C91 = tables.tab_Cost_Model_L_PSH_C90_C91
tab_Cost_Model_L_PSH_C97_C105 = tables.tab_Cost_Model_L_PSH_C97_C105
tab_Cost_Model_L_PSH_C108_C111 = tables.tab_Cost_Model_L_PSH_C108_C111
tab_Cost_Curves_S_PSH_F87_H93 = tables.tab_Cost_Curves_S_PSH_F87_H93
tab_Cost_Curves_S_PSH_J95_J96 = tables.tab_Cost_Curves_S_PSH_J95_J96
tab_Cost_Curves_S_PSH_K290_K292 = tables.tab_Cost_Curves_S_PSH_K290_K292
tab_Market_Adj_Factors_I6_I58 = tables.tab_Market_Adj_Factors_I6_I58
tab_Market_Adj_Factors_D30_D43 = tables.tab_Market_Adj_Factors_D30_D43
tab_Market_Adj_Factors_C45_C48 = tables.tab_Market_Adj_Factors_C45_C48
tab_Land_Value_Data_B44_F44 = tables.tab_Land_Value_Data_B44_F44
tab_Land_Value_Data_B8_F8 = tables.tab_Land_Value_Data_B8_F8
tab_Land_Value_Data_B17_F17 = tables.tab_Land_Value_Data_B17_F17
tab_Locational_Adj_Factors_A3_B55 = tables.tab_Locational_Adj_Factors_A3_B55
tab_Land_Value_Data_B38_F38 = tables.tab_Land_Value_Data_B38_F38
tab_Land_Value_Data_B12_F12 = tables.tab_Land_Value_Data_B12_F12
tab_Land_Value_Data_B20_F20 = tables.tab_Land_Value_Data_B20_F20
tab_Land_Value_Data_B36_F36 = tables.tab_Land_Value_Data_B36_F36
tab_Land_Value_Data_B4_F4 = tables.tab_Land_Value_Data_B4_F4
tab_Land_Value_Data_B43_F43 = tables.tab_Land_Value_Data_B43_F43
tab_Land_Value_Data_B11_F11 = tables.tab_Land_Value_Data_B11_F11
tab_Land_Value_Data_B25_F25 = tables.tab_Land_Value_Data_B25_F25
tab_Land_Value_Data_B5_F5 = tables.tab_Land_Value_Data_B5_F5
tab_Land_Value_Data_B31_F31 = tables.tab_Land_Value_Data_B31_F31
tab_Land_Value_Data_B46_F46 = tables.tab_Land_Value_Data_B46_F46
tab_Land_Value_Data_B34_F34 = tables.tab_Land_Value_Data_B34_F34
tab_Land_Value_Data_B32_F32 = tables.tab_Land_Value_Data_B32_F32
tab_Land_Value_Data_B18_F18 = tables.tab_Land_Value_Data_B18_F18
tab_Land_Value_Data_B42_F42 = tables.tab_Land_Value_Data_B42_F42
tab_Land_Value_Data_B49_F49 = tables.tab_Land_Value_Data_B49_F49
tab_Land_Value_Data_B15_F15 = tables.tab_Land_Value_Data_B15_F15
tab_Land_Value_Data_B14_F14 = tables.tab_Land_Value_Data_B14_F14
tab_Land_Value_Data_B35_F35 = tables.tab_Land_Value_Data_B35_F35
tab_Land_Value_Data_B7_F7 = tables.tab_Land_Value_Data_B7_F7
tab_Land_Value_Data_B41_F41 = tables.tab_Land_Value_Data_B41_F41
tab_Land_Value_Data_B33_F33 = tables.tab_Land_Value_Data_B33_F33
tab_Land_Value_Data_B45_F45 = tables.tab_Land_Value_Data_B45_F45
tab_Land_Value_Data_B22_F22 = tables.tab_Land_Value_Data_B22_F22
tab_Land_Value_Data_B47_F47 = tables.tab_Land_Value_Data_B47_F47
tab_Land_Value_Data_B50_F50 = tables.tab_Land_Value_Data_B50_F50
tab_Land_Value_Data_B9_F9 = tables.tab_Land_Value_Data_B9_F9
tab_Land_Value_Data_B30_F30 = tables.tab_Land_Value_Data_B30_F30
tab_Land_Value_Data_B27_F27 = tables.tab_Land_Value_Data_B27_F27
tab_Land_Value_Data_B16_F16 = tables.tab_Land_Value_Data_B16_F16
tab_Land_Value_Data_B28_F28 = tables.tab_Land_Value_Data_B28_F28
tab_Land_Value_Data_B21_F21 = tables.tab_Land_Value_Data_B21_F21
tab_Land_Value_Data_B48_F48 = tables.tab_Land_Value_Data_B48_F48
tab_Land_Value_Data_B37_F37 = tables.tab_Land_Value_Data_B37_F37
tab_Land_Value_Data_B39_F39 = tables.tab_Land_Value_Data_B39_F39
tab_Land_Value_Data_B6_F6 = tables.tab_Land_Value_Data_B6_F6
tab_Land_Value_Data_B51_F51 = tables.tab_Land_Value_Data_B51_F51
tab_Land_Value_Data_B23_F23 = tables.tab_Land_Value_Data_B23_F23
tab_Land_Value_Data_B10_F10 = tables.tab_Land_Value_Data_B10_F10
tab_Cost_Curves_L_PSH_W241_X243 = tables.tab_Cost_Curves_L_PSH_W241_X243
tab_Land_Value_Data_B29_F29 = tables.tab_Land_Value_Data_B29_F29
tab_Land_Value_Data_B13_F13 = tables.tab_Land_Value_Data_B13_F13
tab_Land_Value_Data_B26_F26 = tables.tab_Land_Value_Data_B26_F26
tab_Land_Value_Data_B3_F3 = tables.tab_Land_Value_Data_B3_F3
tab_Land_Value_Data_B24_F24 = tables.tab_Land_Value_Data_B24_F24
tab_Land_Value_Data_B40_F40 = tables.tab_Land_Value_Data_B40_F40
tab_Land_Value_Data_B19_F19 = tables.tab_Land_Value_Data_B19_F19
tab_Cost_Curves_L_PSH_H276_I284 = tables.tab_Cost_Curves_L_PSH_H276_I284
# Level 0


# Level 1
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_C80_C86[6, "C"])]
tab_Cost_Model_L_PSH_C80_C86[2, "C"] = inputs.Upper_Reservoir_Area_Cost_Model_L_PSH_C11 * inputs.Avg_Max_Upper_Reservoir_Depth_Cost_Model_L_PSH_C10 # Cost Model L-PSH C81 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_C80_C86[2, "C"], 19291) # "Cost Model L-PSH!C81"
# Used in 3 places: [StandardStatement(lhs = s_Cost_Curves_L_PSH_AA130), StandardStatement(lhs = s_Cost_Model_L_PSH_J23), StandardStatement(lhs = Mean_Gross_Head__Cost_Model_L_PSH_C89)]
# =C69*C14
Min_Gross_Head__Cost_Model_L_PSH_C88 = inputs.Hmin_per_Hmax__Cost_Model_L_PSH_C69 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 # Cost Model L-PSH C88
@assert xl_compare(Min_Gross_Head__Cost_Model_L_PSH_C88, 1092) # "Cost Model L-PSH!C88"
# Used in 17 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[1, "L"]), TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[6, "L"]), StandardStatement(lhs = s_Cost_Model_L_PSH_L14), TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[2, "L"]), ..., FunctionStatement(lhs = s_Cost_Model_L_PSH_N39)]
tab_Market_Adj_Factors_C45_C48[1, "C"] = inputs.s_Market_Adj_Factors_H58 / inputs.s_Market_Adj_Factors_H24 # Market Adj Factors C45 Row: 1
@assert xl_compare(tab_Market_Adj_Factors_C45_C48[1, "C"], 2.4063214260550976) # "Market Adj Factors!C45"


# Level 2
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =740.12*('Cost Model L-PSH'!C14)^-0.298
Est__dollar_per_kW_Cost_Curves_L_PSH_B11 = 740.12 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.298))) # Cost Curves L-PSH B11
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_B11, 82.7467623110216) # "Cost Curves L-PSH!B11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1229.6*('Cost Model L-PSH'!C14)^-0.405
Est__dollar_per_kW_Cost_Curves_L_PSH_B53 = 1229.6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.405))) # Cost Curves L-PSH B53
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_B53, 62.596260921003555) # "Cost Curves L-PSH!B53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =760.42*('Cost Model L-PSH'!C14)^-0.328
s_Cost_Curves_L_PSH_C11 = 760.42 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.328))) # Cost Curves L-PSH C11
@assert xl_compare(s_Cost_Curves_L_PSH_C11, 68.1881114906183) # "Cost Curves L-PSH!C11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1113.2*('Cost Model L-PSH'!C14)^-0.412
s_Cost_Curves_L_PSH_C53 = 1113.2 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.412))) # Cost Curves L-PSH C53
@assert xl_compare(s_Cost_Curves_L_PSH_C53, 53.827704941124054) # "Cost Curves L-PSH!C53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =645.72*('Cost Model L-PSH'!C14)^-0.32
s_Cost_Curves_L_PSH_D11 = 645.72 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.32))) # Cost Curves L-PSH D11
@assert xl_compare(s_Cost_Curves_L_PSH_D11, 61.41074508407125) # "Cost Curves L-PSH!D11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1126.8*('Cost Model L-PSH'!C14)^-0.431
s_Cost_Curves_L_PSH_D53 = 1126.8 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.431))) # Cost Curves L-PSH D53
@assert xl_compare(s_Cost_Curves_L_PSH_D53, 47.38164520745838) # "Cost Curves L-PSH!D53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =640.28*('Cost Model L-PSH'!C14)^-0.337
s_Cost_Curves_L_PSH_E11 = 640.28 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.337))) # Cost Curves L-PSH E11
@assert xl_compare(s_Cost_Curves_L_PSH_E11, 53.73867407370265) # "Cost Curves L-PSH!E11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1195.5*('Cost Model L-PSH'!C14)^-0.454
s_Cost_Curves_L_PSH_E53 = 1195.5 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.454))) # Cost Curves L-PSH E53
@assert xl_compare(s_Cost_Curves_L_PSH_E53, 42.44934689608617) # "Cost Curves L-PSH!E53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1096.7*('Cost Model L-PSH'!C14)^-0.319
s_Cost_Curves_L_PSH_F11 = 1096.7 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.319))) # Cost Curves L-PSH F11
@assert xl_compare(s_Cost_Curves_L_PSH_F11, 105.0705720118597) # "Cost Curves L-PSH!F11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1435.5*('Cost Model L-PSH'!C14)^-0.392
s_Cost_Curves_L_PSH_F53 = 1435.5 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.392))) # Cost Curves L-PSH F53
@assert xl_compare(s_Cost_Curves_L_PSH_F53, 80.40783231276411) # "Cost Curves L-PSH!F53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =878.1*('Cost Model L-PSH'!C14)^-0.313
s_Cost_Curves_L_PSH_G11 = 878.1 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.313))) # Cost Curves L-PSH G11
@assert xl_compare(s_Cost_Curves_L_PSH_G11, 87.921679982121) # "Cost Curves L-PSH!G11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1258*('Cost Model L-PSH'!C14)^-0.396
s_Cost_Curves_L_PSH_G53 = 1258.0 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.396))) # Cost Curves L-PSH G53
@assert xl_compare(s_Cost_Curves_L_PSH_G53, 68.42318721566281) # "Cost Curves L-PSH!G53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =926.99*('Cost Model L-PSH'!C14)^-0.337
s_Cost_Curves_L_PSH_H11 = 926.99 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.337))) # Cost Curves L-PSH H11
@assert xl_compare(s_Cost_Curves_L_PSH_H11, 77.80223258509031) # "Cost Curves L-PSH!H11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1375*('Cost Model L-PSH'!C14)^-0.424
s_Cost_Curves_L_PSH_H53 = 1375.0 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.424))) # Cost Curves L-PSH H53
@assert xl_compare(s_Cost_Curves_L_PSH_H53, 60.87204168048313) # "Cost Curves L-PSH!H53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =911.22*('Cost Model L-PSH'!C14)^-0.352
s_Cost_Curves_L_PSH_I11 = 911.22 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.352))) # Cost Curves L-PSH I11
@assert xl_compare(s_Cost_Curves_L_PSH_I11, 68.49255777082976) # "Cost Curves L-PSH!I11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1131*('Cost Model L-PSH'!C14)^-0.41
s_Cost_Curves_L_PSH_I53 = 1131.0 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.41))) # Cost Curves L-PSH I53
@assert xl_compare(s_Cost_Curves_L_PSH_I53, 55.498535055978614) # "Cost Curves L-PSH!I53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2032.8*('Cost Model L-PSH'!C14)^-0.532
Est__dollar_per_kW_Cost_Curves_L_PSH_L11 = 2032.8 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.532))) # Cost Curves L-PSH L11
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_L11, 40.67732316085264) # "Cost Curves L-PSH!L11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2163.3*('Cost Model L-PSH'!C14)^-0.512
Est__dollar_per_kW_Cost_Curves_L_PSH_L53 = 2163.3 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.512))) # Cost Curves L-PSH L53
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_L53, 50.14607615964969) # "Cost Curves L-PSH!L53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2231.2*('Cost Model L-PSH'!C14)^-0.565
s_Cost_Curves_L_PSH_M11 = 2231.2 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.565))) # Cost Curves L-PSH M11
@assert xl_compare(s_Cost_Curves_L_PSH_M11, 35.02862606158792) # "Cost Curves L-PSH!M11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1725.2*('Cost Model L-PSH'!C14)^-0.501
s_Cost_Curves_L_PSH_M53 = 1725.2 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.501))) # Cost Curves L-PSH M53
@assert xl_compare(s_Cost_Curves_L_PSH_M53, 43.35947605914813) # "Cost Curves L-PSH!M53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2553*('Cost Model L-PSH'!C14)^-0.604
s_Cost_Curves_L_PSH_N11 = 2553.0 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.604))) # Cost Curves L-PSH N11
@assert xl_compare(s_Cost_Curves_L_PSH_N11, 30.08870922844283) # "Cost Curves L-PSH!N11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1735.2*('Cost Model L-PSH'!C14)^-0.52
s_Cost_Curves_L_PSH_N53 = 1735.2 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.52))) # Cost Curves L-PSH N53
@assert xl_compare(s_Cost_Curves_L_PSH_N53, 37.92492650716327) # "Cost Curves L-PSH!N53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2504.2*('Cost Model L-PSH'!C14)^-0.615
s_Cost_Curves_L_PSH_O11 = 2504.2 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.615))) # Cost Curves L-PSH O11
@assert xl_compare(s_Cost_Curves_L_PSH_O11, 27.220580536976755) # "Cost Curves L-PSH!O11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =1578.2*('Cost Model L-PSH'!C14)^-0.524
s_Cost_Curves_L_PSH_O53 = 1578.2 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.524))) # Cost Curves L-PSH O53
@assert xl_compare(s_Cost_Curves_L_PSH_O53, 33.49382520354347) # "Cost Curves L-PSH!O53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2613*('Cost Model L-PSH'!C14)^-0.533
s_Cost_Curves_L_PSH_P11 = 2613.0 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.533))) # Cost Curves L-PSH P11
@assert xl_compare(s_Cost_Curves_L_PSH_P11, 51.90437893161324) # "Cost Curves L-PSH!P11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2315.8*('Cost Model L-PSH'!C14)^-0.49
s_Cost_Curves_L_PSH_P53 = 2315.8 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.49))) # Cost Curves L-PSH P53
@assert xl_compare(s_Cost_Curves_L_PSH_P53, 63.105904575527866) # "Cost Curves L-PSH!P53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2154.7*('Cost Model L-PSH'!C14)^-0.53
s_Cost_Curves_L_PSH_Q11 = 2154.7 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.53))) # Cost Curves L-PSH Q11
@assert xl_compare(s_Cost_Curves_L_PSH_Q11, 43.75531095009447) # "Cost Curves L-PSH!Q11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2194.6*('Cost Model L-PSH'!C14)^-0.505
s_Cost_Curves_L_PSH_Q53 = 2194.6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.505))) # Cost Curves L-PSH Q53
@assert xl_compare(s_Cost_Curves_L_PSH_Q53, 53.55838251835405) # "Cost Curves L-PSH!Q53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2380.6*('Cost Model L-PSH'!C14)^-0.559
s_Cost_Curves_L_PSH_R11 = 2380.6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.559))) # Cost Curves L-PSH R11
@assert xl_compare(s_Cost_Curves_L_PSH_R11, 39.0597787354559) # "Cost Curves L-PSH!R11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2353.1*('Cost Model L-PSH'!C14)^-0.53
s_Cost_Curves_L_PSH_R53 = 2353.1 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.53))) # Cost Curves L-PSH R53
@assert xl_compare(s_Cost_Curves_L_PSH_R53, 47.784202996550476) # "Cost Curves L-PSH!R53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =4996.8*('Cost Model L-PSH'!C14)^-0.687
s_Cost_Curves_L_PSH_S11 = 4996.8 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.687))) # Cost Curves L-PSH S11
@assert xl_compare(s_Cost_Curves_L_PSH_S11, 31.9900830685423) # "Cost Curves L-PSH!S11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =2162.4*('Cost Model L-PSH'!C14)^-0.533
s_Cost_Curves_L_PSH_S53 = 2162.4 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ ((-1 * 0.533))) # Cost Curves L-PSH S53
@assert xl_compare(s_Cost_Curves_L_PSH_S53, 42.953704172108864) # "Cost Curves L-PSH!S53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000005*('Cost Model L-PSH'!C14)^2 - 0.0009*('Cost Model L-PSH'!C14) + 83.095
Est__dollar_per_kW_Cost_Curves_L_PSH_V11 = 5.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0009 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 83.095 # Cost Curves L-PSH V11
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_V11, 93.859) # "Cost Curves L-PSH!V11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000007*('Cost Model L-PSH'!C14)^2 - 0.0146*('Cost Model L-PSH'!C14) + 67.574
Est__dollar_per_kW_Cost_Curves_L_PSH_V53 = 7.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0146 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 67.574 # Cost Curves L-PSH V53
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_V53, 61.8332) # "Cost Curves L-PSH!V53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000004*('Cost Model L-PSH'!C14)^2 + 0.0004*('Cost Model L-PSH'!C14) + 75.715
s_Cost_Curves_L_PSH_W11 = 4.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0004 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 75.715 # Cost Curves L-PSH W11
@assert xl_compare(s_Cost_Curves_L_PSH_W11, 86.0734) # "Cost Curves L-PSH!W11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000007*('Cost Model L-PSH'!C14)^2 - 0.0117*('Cost Model L-PSH'!C14) + 59.042
s_Cost_Curves_L_PSH_W53 = 7.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0117 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 59.042 # Cost Curves L-PSH W53
@assert xl_compare(s_Cost_Curves_L_PSH_W53, 57.8252) # "Cost Curves L-PSH!W53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000004*('Cost Model L-PSH'!C14)^2 + 0.0021*('Cost Model L-PSH'!C14) + 71.437
s_Cost_Curves_L_PSH_X11 = 4.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0021 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 71.437 # Cost Curves L-PSH X11
@assert xl_compare(s_Cost_Curves_L_PSH_X11, 84.4474) # "Cost Curves L-PSH!X11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000005*('Cost Model L-PSH'!C14)^2 - 0.0065*('Cost Model L-PSH'!C14) + 54.377
s_Cost_Curves_L_PSH_X53 = 5.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0065 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 54.377 # Cost Curves L-PSH X53
@assert xl_compare(s_Cost_Curves_L_PSH_X53, 56.405) # "Cost Curves L-PSH!X53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.0000006*('Cost Model L-PSH'!C14)^2 + 0.0109*('Cost Model L-PSH'!C14) + 62.383
s_Cost_Curves_L_PSH_Y11 = 6.0e-7 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0109 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 62.383 # Cost Curves L-PSH Y11
@assert xl_compare(s_Cost_Curves_L_PSH_Y11, 80.84716) # "Cost Curves L-PSH!Y11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000003*('Cost Model L-PSH'!C14)^2 - 0.0018*('Cost Model L-PSH'!C14) + 48.04
s_Cost_Curves_L_PSH_Y53 = 3.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0018 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 48.04 # Cost Curves L-PSH Y53
@assert xl_compare(s_Cost_Curves_L_PSH_Y53, 52.5328) # "Cost Curves L-PSH!Y53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000003*('Cost Model L-PSH'!C14)^2 + 0.0062*('Cost Model L-PSH'!C14) + 109.48
s_Cost_Curves_L_PSH_Z11 = 3.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0062 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 109.48 # Cost Curves L-PSH Z11
@assert xl_compare(s_Cost_Curves_L_PSH_Z11, 126.4528) # "Cost Curves L-PSH!Z11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000004*('Cost Model L-PSH'!C14)^2 - 0.0039*('Cost Model L-PSH'!C14) + 90.761
s_Cost_Curves_L_PSH_Z53 = 4.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0039 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 90.761 # Cost Curves L-PSH Z53
@assert xl_compare(s_Cost_Curves_L_PSH_Z53, 94.4114) # "Cost Curves L-PSH!Z53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000001*('Cost Model L-PSH'!C14)^2 + 0.0099*('Cost Model L-PSH'!C14) + 90.767
s_Cost_Curves_L_PSH_AA11 = 1.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0099 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 90.767 # Cost Curves L-PSH AA11
@assert xl_compare(s_Cost_Curves_L_PSH_AA11, 108.6446) # "Cost Curves L-PSH!AA11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000002*('Cost Model L-PSH'!C14)^2 + 0.0008*('Cost Model L-PSH'!C14) + 75.02
s_Cost_Curves_L_PSH_AA53 = 2.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0008 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 75.02 # Cost Curves L-PSH AA53
@assert xl_compare(s_Cost_Curves_L_PSH_AA53, 81.1352) # "Cost Curves L-PSH!AA53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000002*('Cost Model L-PSH'!C14)^2 + 0.0073*('Cost Model L-PSH'!C14) + 84.676
s_Cost_Curves_L_PSH_AB11 = 2.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0073 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 84.676 # Cost Curves L-PSH AB11
@assert xl_compare(s_Cost_Curves_L_PSH_AB11, 100.9312) # "Cost Curves L-PSH!AB11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000002*('Cost Model L-PSH'!C14)^2 + 0.0002*('Cost Model L-PSH'!C14) + 70.655
s_Cost_Curves_L_PSH_AB53 = 2.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0002 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 70.655 # Cost Curves L-PSH AB53
@assert xl_compare(s_Cost_Curves_L_PSH_AB53, 75.8342) # "Cost Curves L-PSH!AB53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =-0.0000001*('Cost Model L-PSH'!C14)^2 + 0.0123*('Cost Model L-PSH'!C14) + 73.167
s_Cost_Curves_L_PSH_AC11 = (-1 * 1.0e-7) * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.0123 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 73.167 # Cost Curves L-PSH AC11
@assert xl_compare(s_Cost_Curves_L_PSH_AC11, 92.11164) # "Cost Curves L-PSH!AC11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000002*('Cost Model L-PSH'!C14)^2 + 0.002*('Cost Model L-PSH'!C14)+ 62.446
s_Cost_Curves_L_PSH_AC53 = 2.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) + 0.002 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 62.446 # Cost Curves L-PSH AC53
@assert xl_compare(s_Cost_Curves_L_PSH_AC53, 70.4332) # "Cost Curves L-PSH!AC53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000006*('Cost Model L-PSH'!C14)^2 - 0.0183*('Cost Model L-PSH'!C14) + 43.547
Est__dollar_per_kW_Cost_Curves_L_PSH_AF11 = 6.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0183 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 43.547 # Cost Curves L-PSH AF11
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_AF11, 29.600599999999996) # "Cost Curves L-PSH!AF11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000006*('Cost Model L-PSH'!C14)^2 - 0.0159*('Cost Model L-PSH'!C14) + 53.593
Est__dollar_per_kW_Cost_Curves_L_PSH_AF53 = 6.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0159 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 53.593 # Cost Curves L-PSH AF53
@assert xl_compare(Est__dollar_per_kW_Cost_Curves_L_PSH_AF53, 43.390600000000006) # "Cost Curves L-PSH!AF53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000007*('Cost Model L-PSH'!C14)^2 - 0.0198*('Cost Model L-PSH'!C14) + 41.489
s_Cost_Curves_L_PSH_AG11 = 7.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0198 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 41.489 # Cost Curves L-PSH AG11
@assert xl_compare(s_Cost_Curves_L_PSH_AG11, 27.636199999999995) # "Cost Curves L-PSH!AG11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000006*('Cost Model L-PSH'!C14)^2 - 0.0157*('Cost Model L-PSH'!C14) + 47.989
s_Cost_Curves_L_PSH_AG53 = 6.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0157 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 47.989 # Cost Curves L-PSH AG53
@assert xl_compare(s_Cost_Curves_L_PSH_AG53, 38.098600000000005) # "Cost Curves L-PSH!AG53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000006*('Cost Model L-PSH'!C14)^2 - 0.0162*('Cost Model L-PSH'!C14) + 36.672
s_Cost_Curves_L_PSH_AH11 = 6.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0162 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 36.672 # Cost Curves L-PSH AH11
@assert xl_compare(s_Cost_Curves_L_PSH_AH11, 26.0016) # "Cost Curves L-PSH!AH11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000005*('Cost Model L-PSH'!C14)^2 - 0.0116*('Cost Model L-PSH'!C14)+ 43.522
s_Cost_Curves_L_PSH_AH53 = 5.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0116 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 43.522 # Cost Curves L-PSH AH53
@assert xl_compare(s_Cost_Curves_L_PSH_AH53, 37.594) # "Cost Curves L-PSH!AH53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000005*('Cost Model L-PSH'!C14)^2 - 0.0128*('Cost Model L-PSH'!C14) + 33.573
s_Cost_Curves_L_PSH_AI11 = 5.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0128 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 33.573 # Cost Curves L-PSH AI11
@assert xl_compare(s_Cost_Curves_L_PSH_AI11, 25.773000000000003) # "Cost Curves L-PSH!AI11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000003*('Cost Model L-PSH'!C14)^2 - 0.0078*('Cost Model L-PSH'!C14) + 39.647
s_Cost_Curves_L_PSH_AI53 = 3.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0078 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 39.647 # Cost Curves L-PSH AI53
@assert xl_compare(s_Cost_Curves_L_PSH_AI53, 34.7798) # "Cost Curves L-PSH!AI53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.00001*('Cost Model L-PSH'!C14)^2 - 0.0308*('Cost Model L-PSH'!C14) + 67.023
s_Cost_Curves_L_PSH_AJ11 = 1.0e-5 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0308 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 67.023 # Cost Curves L-PSH AJ11
@assert xl_compare(s_Cost_Curves_L_PSH_AJ11, 43.31099999999999) # "Cost Curves L-PSH!AJ11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000008*('Cost Model L-PSH'!C14)^2 - 0.0241*('Cost Model L-PSH'!C14)+ 80.261
s_Cost_Curves_L_PSH_AJ53 = 8.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0241 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 80.261 # Cost Curves L-PSH AJ53
@assert xl_compare(s_Cost_Curves_L_PSH_AJ53, 62.133799999999994) # "Cost Curves L-PSH!AJ53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000007*('Cost Model L-PSH'!C14)^2 - 0.0207*('Cost Model L-PSH'!C14) + 54.238
s_Cost_Curves_L_PSH_AK11 = 7.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0207 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 54.238 # Cost Curves L-PSH AK11
@assert xl_compare(s_Cost_Curves_L_PSH_AK11, 38.9812) # "Cost Curves L-PSH!AK11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000005*('Cost Model L-PSH'!C14)^2 - 0.0155*('Cost Model L-PSH'!C14) + 65.884
s_Cost_Curves_L_PSH_AK53 = 5.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0155 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 65.884 # Cost Curves L-PSH AK53
@assert xl_compare(s_Cost_Curves_L_PSH_AK53, 53.872) # "Cost Curves L-PSH!AK53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000007*('Cost Model L-PSH'!C14)^2 - 0.0205*('Cost Model L-PSH'!C14) + 50.886
s_Cost_Curves_L_PSH_AL11 = 7.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0205 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 50.886 # Cost Curves L-PSH AL11
@assert xl_compare(s_Cost_Curves_L_PSH_AL11, 35.9412) # "Cost Curves L-PSH!AL11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000006*('Cost Model L-PSH'!C14)^2 - 0.0165*('Cost Model L-PSH'!C14) + 61.35
s_Cost_Curves_L_PSH_AL53 = 6.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0165 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 61.35 # Cost Curves L-PSH AL53
@assert xl_compare(s_Cost_Curves_L_PSH_AL53, 50.211600000000004) # "Cost Curves L-PSH!AL53"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000005*('Cost Model L-PSH'!C14)^2 - 0.0146*('Cost Model L-PSH'!C14) + 43.361
s_Cost_Curves_L_PSH_AM11 = 5.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0146 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 43.361 # Cost Curves L-PSH AM11
@assert xl_compare(s_Cost_Curves_L_PSH_AM11, 32.753) # "Cost Curves L-PSH!AM11"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =0.000004*('Cost Model L-PSH'!C14)^2 - 0.0099*('Cost Model L-PSH'!C14) + 52.44
s_Cost_Curves_L_PSH_AM53 = 4.0e-6 * ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) ^ (2.0)) - 0.0099 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 52.44 # Cost Curves L-PSH AM53
@assert xl_compare(s_Cost_Curves_L_PSH_AM53, 46.730399999999996) # "Cost Curves L-PSH!AM53"
# Used in 2 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10)), FunctionStatement(lhs = s_Cost_Model_L_PSH_J28)]
tab_Cost_Model_L_PSH_C80_C86[6, "C"] = inputs.Active_Storage__Cost_Model_L_PSH_C68 * tab_Cost_Model_L_PSH_C80_C86[2, "C"] # Cost Model L-PSH C85 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_C80_C86[6, "C"], 16397.35) # "Cost Model L-PSH!C85"
# Used in 9 places: [StandardStatement(lhs = Ft_Cost_Model_L_PSH_I24), StandardStatement(lhs = Ft_Cost_Model_L_PSH_I20), StandardStatement(lhs = Ft_Cost_Model_L_PSH_I22), StandardStatement(lhs = Ft_Cost_Model_L_PSH_I25), ..., FunctionStatement(lhs = s_Cost_Model_L_PSH_J30)]
# =(C88+C14)/2
Mean_Gross_Head__Cost_Model_L_PSH_C89 = (Min_Gross_Head__Cost_Model_L_PSH_C88 + inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14) / 2.0 # Cost Model L-PSH C89
@assert xl_compare(Mean_Gross_Head__Cost_Model_L_PSH_C89, 1326) # "Cost Model L-PSH!C89"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =VLOOKUP(C9,'Locational Adj Factors'!$A$3:$B$55,2,FALSE)/100
s_Cost_Model_L_PSH_K10 = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K10
@assert xl_compare(s_Cost_Model_L_PSH_K10, 1.0) # "Cost Model L-PSH!K10"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# =IF(C51="Yes",'Market Adj Factors'!C45,1)
s_Cost_Model_L_PSH_L10 = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L10
@assert xl_compare(s_Cost_Model_L_PSH_L10, 2.4063214260550976) # "Cost Model L-PSH!L10"
# Used in 1 places: [GroupedStatement(StandardStatement(lhs = Mean_Gen_Discharge__Cost_Model_L_PSH_C87), TableStatement(lhs = tab_Cost_Model_L_PSH_C90_C91[2, "C"]), StandardStatement(lhs = Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92), StandardStatement(lhs = No_Tunnels__Cost_Model_L_PSH_C93), StandardStatement(lhs = Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[3, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[6, "C"]), TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[9, "C"]), StandardStatement(lhs = No_Units__Cost_Model_L_PSH_C106), StandardStatement(lhs = Unit_Rating__Cost_Model_L_PSH_C107), StandardStatement(lhs = s_Cost_Model_L_PSH_J10), StandardStatement(lhs = s_Cost_Model_L_PSH_N10))]
# ='Market Adj Factors'!C30
s_Cost_Model_L_PSH_M10 = inputs.Power_station # Cost Model L-PSH M10
@assert xl_compare(s_Cost_Model_L_PSH_M10, 1.3) # "Cost Model L-PSH!M10"


# Level 3
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_C108_C111[1, "C"])]
# =0.0074*'Cost Model L-PSH'!C14 + 16.512
s_Cost_Curves_L_PSH_V130 = 0.0074 * inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 + 16.512 # Cost Curves L-PSH V130
@assert xl_compare(s_Cost_Curves_L_PSH_V130, 28.056) # "Cost Curves L-PSH!V130"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_C108_C111[2, "C"])]
# =0.0046*('Cost Model L-PSH'!C14-'Cost Model L-PSH'!C88) + 7.516
s_Cost_Curves_L_PSH_AA130 = 0.0046 * (inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - Min_Gross_Head__Cost_Model_L_PSH_C88) + 7.516 # Cost Curves L-PSH AA130
@assert xl_compare(s_Cost_Curves_L_PSH_AA130, 9.668800000000001) # "Cost Curves L-PSH!AA130"
# Used in 31 places: [StandardStatement(lhs = s_Cost_Curves_L_PSH_AB95), StandardStatement(lhs = Ft_Cost_Model_L_PSH_I24), TableStatement(lhs = tab_Cost_Model_L_PSH_I13_I14[2, "I"]), StandardStatement(lhs = s_Cost_Curves_L_PSH_Z95), ..., FunctionStatement(lhs = s_Cost_Model_L_PSH_Q42)]
# Group of 12 statements
begin
# =C85*C26/C21/C28
Mean_Gen_Discharge__Cost_Model_L_PSH_C87 = tab_Cost_Model_L_PSH_C80_C86[6, "C"] * inputs.Ac_Ft_to_Cu_Ft_Cost_Model_L_PSH_C26 / inputs.Generation_Time_Cost_Model_L_PSH_C21 / inputs.Hr_to_Sec_Cost_Model_L_PSH_C28 # Cost Model L-PSH C87
@assert xl_compare(Mean_Gen_Discharge__Cost_Model_L_PSH_C87, 10724.728622597597) # "Cost Model L-PSH!C87"
tab_Cost_Model_L_PSH_C90_C91[2, "C"] = ((inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 / Mean_Gross_Head__Cost_Model_L_PSH_C89) ^ (0.5)) * Mean_Gen_Discharge__Cost_Model_L_PSH_C87 # Cost Model L-PSH C91 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_C90_C91[2, "C"], 11632.601450404713) # "Cost Model L-PSH!C91"
# =((C91/C72)*4/PI())^0.5
Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92 = ((tab_Cost_Model_L_PSH_C90_C91[2, "C"] / inputs.Max_Tunnel_Velocity__Cost_Model_L_PSH_C72 * 4.0 / π) ^ (0.5)) # Cost Model L-PSH C92
@assert xl_compare(Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, 25.3763739613452) # "Cost Model L-PSH!C92"
# =IF(C92<C73,1,2)
No_Tunnels__Cost_Model_L_PSH_C93 = (xl_lt(Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, inputs.Max_Tunnel_Dia__Cost_Model_L_PSH_C73) ? 1.0 : 2.0) # Cost Model L-PSH C93
@assert xl_compare(No_Tunnels__Cost_Model_L_PSH_C93, 1) # "Cost Model L-PSH!C93"
# =((C91/C93/C72)*4/PI())^0.5
Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 = ((tab_Cost_Model_L_PSH_C90_C91[2, "C"] / No_Tunnels__Cost_Model_L_PSH_C93 / inputs.Max_Tunnel_Velocity__Cost_Model_L_PSH_C72 * 4.0 / π) ^ (0.5)) # Cost Model L-PSH C94
@assert xl_compare(Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, 25.3763739613452) # "Cost Model L-PSH!C94"
tab_Cost_Model_L_PSH_C97_C105[3, "C"] = 4.73 * ((tab_Cost_Model_L_PSH_C90_C91[2, "C"]) ^ (1.85)) / ((inputs.Hazen_Williams_C__Cost_Model_L_PSH_C70) ^ (1.85)) / ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (4.87)) * inputs.Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15 # Cost Model L-PSH C99 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[3, "C"], 79.36411576234892) # "Cost Model L-PSH!C99"
tab_Cost_Model_L_PSH_C97_C105[6, "C"] = inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - tab_Cost_Model_L_PSH_C97_C105[3, "C"] # Cost Model L-PSH C102 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[6, "C"], 1480.635884237651) # "Cost Model L-PSH!C102"
tab_Cost_Model_L_PSH_C97_C105[9, "C"] = tab_Cost_Model_L_PSH_C90_C91[2, "C"] * tab_Cost_Model_L_PSH_C97_C105[6, "C"] * inputs.P_T_Efficiency__Cost_Model_L_PSH_C71 * inputs.acceleration_of_gravity_metric_Cost_Model_L_PSH_C30 * inputs.density_of_water_metric_Cost_Model_L_PSH_C31 / inputs.MW_to_W_Cost_Model_L_PSH_C29 * ((1.0 / inputs.Meter_to_Feet_Cost_Model_L_PSH_C24) ^ (4.0)) # Cost Model L-PSH C105 Row: 9
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[9, "C"], 1283.3248207034676) # "Cost Model L-PSH!C105"
# =IF(C105/C74<=C75,C75,ROUNDUP(C105/C74,0))
No_Units__Cost_Model_L_PSH_C106 = (xl_leq(tab_Cost_Model_L_PSH_C97_C105[9, "C"] / inputs.Max_Unit_Capacity__Cost_Model_L_PSH_C74, inputs.Min_No_Units__Cost_Model_L_PSH_C75) ? inputs.Min_No_Units__Cost_Model_L_PSH_C75 : round(tab_Cost_Model_L_PSH_C97_C105[9, "C"] / inputs.Max_Unit_Capacity__Cost_Model_L_PSH_C74, RoundFromZero, digits=Int(0.0))) # Cost Model L-PSH C106
@assert xl_compare(No_Units__Cost_Model_L_PSH_C106, 4) # "Cost Model L-PSH!C106"
# =C105/C106
Unit_Rating__Cost_Model_L_PSH_C107 = tab_Cost_Model_L_PSH_C97_C105[9, "C"] / No_Units__Cost_Model_L_PSH_C106 # Cost Model L-PSH C107
@assert xl_compare(Unit_Rating__Cost_Model_L_PSH_C107, 320.8312051758669) # "Cost Model L-PSH!C107"
# =IF(P10,P10,IF(AND(C46="Average",C47="Underground"),IF(C107<=80,IF(C106<=2,'Cost Curves L-PSH'!B11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!C11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!D11,IF(C106>4,'Cost Curves L-PSH'!E11)))),IF(AND(C107>80,C107<=125),IF(C106<=2,'Cost Curves L-PSH'!B53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!C53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!D53,IF(C106>4,'Cost Curves L-PSH'!E53)))),IF(AND(C107>125,C107<=225),IF(C106<=2,'Cost Curves L-PSH'!L53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!M53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!N53,IF(C106>4,'Cost Curves L-PSH'!O53)))),IF(C107>225,IF(C106<=2,'Cost Curves L-PSH'!L11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!M11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!N11,IF(C106>4,'Cost Curves L-PSH'!O11)))))))),IF(AND(C46="Adverse",C47="Underground"),IF(C107<=80,IF(C106<=2,'Cost Curves L-PSH'!F11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!G11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!H11,IF(C106>4,'Cost Curves L-PSH'!I11)))),IF(AND(C107>80,C107<=125),IF(C106<=2,'Cost Curves L-PSH'!F53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!G53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!H53,IF(C106>4,'Cost Curves L-PSH'!I53)))),IF(AND(C107>125,C107<=225),IF(C106<=2,'Cost Curves L-PSH'!P53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!Q53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!R53,IF(C106>4,'Cost Curves L-PSH'!S53)))),IF(C107>225,IF(C106<=2,'Cost Curves L-PSH'!P11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!Q11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!R11,IF(C106>4,'Cost Curves L-PSH'!S11)))))))),IF(AND(C46="Average",C47="Surface"),IF(C107<=80,IF(C106<=2,'Cost Curves L-PSH'!V11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!W11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!X11,IF(C106>4,'Cost Curves L-PSH'!Y11)))),IF(AND(C107>80,C107<=125),IF(C106<=2,'Cost Curves L-PSH'!V53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!W53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!X53,IF(C106>4,'Cost Curves L-PSH'!Y53)))),IF(AND(C107>125,C107<=225),IF(C106<=2,'Cost Curves L-PSH'!AF53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AG53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AH53,IF(C106>4,'Cost Curves L-PSH'!AI53)))),IF(C107>225,IF(C106<=2,'Cost Curves L-PSH'!AF11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AG11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AH11,IF(C106>4,'Cost Curves L-PSH'!AI11)))))))),IF(AND(C46="Adverse",C47="Surface"),IF(C107<=80,IF(C106<=2,'Cost Curves L-PSH'!Z11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AA11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AB11,IF(C106>4,'Cost Curves L-PSH'!AC11)))),IF(AND(C107>80,C107<=125),IF(C106<=2,'Cost Curves L-PSH'!Z53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AA53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AB53,IF(C106>4,'Cost Curves L-PSH'!AC53)))),IF(AND(C107>125,C107<=225),IF(C106<=2,'Cost Curves L-PSH'!AJ53,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AK53,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AL53,IF(C106>4,'Cost Curves L-PSH'!AM53)))),IF(C107>225,IF(C106<=2,'Cost Curves L-PSH'!AJ11,IF(AND(C106>2,C106<=3),'Cost Curves L-PSH'!AK11,IF(AND(C106>3,C106<=4),'Cost Curves L-PSH'!AL11,IF(C106>4,'Cost Curves L-PSH'!AM11)))))))))))))
s_Cost_Model_L_PSH_J10 = (xl_logical(inputs.s_Cost_Model_L_PSH_P10) ? inputs.s_Cost_Model_L_PSH_P10 : (all([xl_eq(inputs.Power_Station_Structure_Geology_Cost_Model_L_PSH_C46, "Average"), xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Underground")]) ? (xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 80.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_B11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_C11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_D11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_E11 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 80.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 125.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_B53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_C53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_D53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_E53 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 125.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 225.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_L53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_M53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_N53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_O53 : missing)))) : (xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 225.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_L11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_M11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_N11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_O11 : missing)))) : missing)))) : (all([xl_eq(inputs.Power_Station_Structure_Geology_Cost_Model_L_PSH_C46, "Adverse"), xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Underground")]) ? (xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 80.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_F11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_G11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_H11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_I11 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 80.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 125.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_F53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_G53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_H53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_I53 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 125.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 225.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_P53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_Q53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_R53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_S53 : missing)))) : (xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 225.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_P11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_Q11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_R11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_S11 : missing)))) : missing)))) : (all([xl_eq(inputs.Power_Station_Structure_Geology_Cost_Model_L_PSH_C46, "Average"), xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Surface")]) ? (xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 80.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_V11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_W11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_X11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_Y11 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 80.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 125.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_V53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_W53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_X53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_Y53 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 125.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 225.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_AF53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AG53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AH53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AI53 : missing)))) : (xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 225.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? Est__dollar_per_kW_Cost_Curves_L_PSH_AF11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AG11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AH11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AI11 : missing)))) : missing)))) : (all([xl_eq(inputs.Power_Station_Structure_Geology_Cost_Model_L_PSH_C46, "Adverse"), xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Surface")]) ? (xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 80.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_Z11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AA11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AB11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AC11 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 80.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 125.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_Z53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AA53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AB53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AC53 : missing)))) : (all([xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 125.0), xl_leq(Unit_Rating__Cost_Model_L_PSH_C107, 225.0)]) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_AJ53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AK53 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AL53 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AM53 : missing)))) : (xl_gt(Unit_Rating__Cost_Model_L_PSH_C107, 225.0) ? (xl_leq(No_Units__Cost_Model_L_PSH_C106, 2.0) ? s_Cost_Curves_L_PSH_AJ11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 2.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 3.0)]) ? s_Cost_Curves_L_PSH_AK11 : (all([xl_gt(No_Units__Cost_Model_L_PSH_C106, 3.0), xl_leq(No_Units__Cost_Model_L_PSH_C106, 4.0)]) ? s_Cost_Curves_L_PSH_AL11 : (xl_gt(No_Units__Cost_Model_L_PSH_C106, 4.0) ? s_Cost_Curves_L_PSH_AM11 : missing)))) : missing)))) : missing))))) # Cost Model L-PSH J10
@assert xl_compare(s_Cost_Model_L_PSH_J10, 39.0597787354559) # "Cost Model L-PSH!J10"
# =IF(P10,P10,J10*K10*L10*M10)
s_Cost_Model_L_PSH_N10 = (xl_logical(inputs.s_Cost_Model_L_PSH_P10) ? inputs.s_Cost_Model_L_PSH_P10 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_K10), s_Cost_Model_L_PSH_L10), s_Cost_Model_L_PSH_M10)) # Cost Model L-PSH N10
@assert xl_compare(s_Cost_Model_L_PSH_N10, 122.18749720852847) # "Cost Model L-PSH!N10"
end
@assert xl_compare(Mean_Gen_Discharge__Cost_Model_L_PSH_C87, 10724.728622597597) # "Cost Model L-PSH!C87"
@assert xl_compare(tab_Cost_Model_L_PSH_C90_C91[2, "C"], 11632.601450404713) # "Cost Model L-PSH!C91"
@assert xl_compare(Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, 25.3763739613452) # "Cost Model L-PSH!C92"
@assert xl_compare(No_Tunnels__Cost_Model_L_PSH_C93, 1) # "Cost Model L-PSH!C93"
@assert xl_compare(Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, 25.3763739613452) # "Cost Model L-PSH!C94"
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[3, "C"], 79.36411576234892) # "Cost Model L-PSH!C99"
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[6, "C"], 1480.635884237651) # "Cost Model L-PSH!C102"
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[9, "C"], 1283.3248207034676) # "Cost Model L-PSH!C105"
@assert xl_compare(No_Units__Cost_Model_L_PSH_C106, 4) # "Cost Model L-PSH!C106"
@assert xl_compare(Unit_Rating__Cost_Model_L_PSH_C107, 320.8312051758669) # "Cost Model L-PSH!C107"
@assert xl_compare(s_Cost_Model_L_PSH_J10, 39.0597787354559) # "Cost Model L-PSH!J10"
@assert xl_compare(s_Cost_Model_L_PSH_N10, 122.18749720852847) # "Cost Model L-PSH!N10"



# Level 4
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[5, "C"])]
tab_Cost_Model_L_PSH_C97_C105[2, "C"] = 4.73 * ((Mean_Gen_Discharge__Cost_Model_L_PSH_C87) ^ (1.85)) / ((inputs.Hazen_Williams_C__Cost_Model_L_PSH_C70) ^ (1.85)) / ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (4.87)) * inputs.Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15 # Cost Model L-PSH C98 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[2, "C"], 68.28678845132482) # "Cost Model L-PSH!C98"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_C108_C111[3, "C"])]
tab_Cost_Model_L_PSH_C108_C111[1, "C"] = s_Cost_Curves_L_PSH_V130 # Cost Model L-PSH C108 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_C108_C111[1, "C"], 28.056) # "Cost Model L-PSH!C108"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_C108_C111[4, "C"])]
tab_Cost_Model_L_PSH_C108_C111[2, "C"] = s_Cost_Curves_L_PSH_AA130 # Cost Model L-PSH C109 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_C108_C111[2, "C"], 9.668800000000001) # "Cost Model L-PSH!C109"


# Level 5
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_C97_C105[8, "C"])]
tab_Cost_Model_L_PSH_C97_C105[5, "C"] = Mean_Gross_Head__Cost_Model_L_PSH_C89 - tab_Cost_Model_L_PSH_C97_C105[2, "C"] # Cost Model L-PSH C101 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[5, "C"], 1257.7132115486752) # "Cost Model L-PSH!C101"
# Used in 5 places: [TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "E"]), TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "D"]), TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "C"]), TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "B"]), FunctionStatement(lhs = s_Cost_Model_L_PSH_J25)]
tab_Cost_Model_L_PSH_C108_C111[3, "C"] = ((tab_Cost_Model_L_PSH_C90_C91[2, "C"] / (tab_Cost_Model_L_PSH_C108_C111[1, "C"] * No_Units__Cost_Model_L_PSH_C106) * 4.0 / π) ^ (0.5)) # Cost Model L-PSH C110 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_C108_C111[3, "C"], 11.488163599886885) # "Cost Model L-PSH!C110"
# Used in 4 places: [TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "E"]), TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "D"]), TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "C"]), TableStatement(lhs = tab_Cost_Curves_L_PSH_B131_E132[!, "B"])]
tab_Cost_Model_L_PSH_C108_C111[4, "C"] = ((tab_Cost_Model_L_PSH_C90_C91[2, "C"] / (tab_Cost_Model_L_PSH_C108_C111[2, "C"] * No_Units__Cost_Model_L_PSH_C106) * 4.0 / π) ^ (0.5)) # Cost Model L-PSH C111 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_C108_C111[4, "C"], 19.569385997752452) # "Cost Model L-PSH!C111"
# Used in 3 places: [StandardStatement(lhs = Ft_Cost_Model_L_PSH_I24), TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[3, "Q"]), StandardStatement(lhs = Ft_Cost_Model_L_PSH_I20)]
# =IF(G22="Yes",IF(C48="Underground",IF(O22,O22,C89*0.25),0))*$C$106
Ft_Cost_Model_L_PSH_I22 = xl_mul((xl_eq(inputs.Penstock_Tunnels_Cost_Model_L_PSH_G22, "Yes") ? (xl_eq(inputs.Penstock_Cost_Model_L_PSH_C48, "Underground") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O22) ? inputs.s_Cost_Model_L_PSH_O22 : Mean_Gross_Head__Cost_Model_L_PSH_C89 * 0.25) : 0.0) : missing), No_Units__Cost_Model_L_PSH_C106) # Cost Model L-PSH I22
@assert xl_compare(Ft_Cost_Model_L_PSH_I22, 1326) # "Cost Model L-PSH!I22"
# Used in 3 places: [StandardStatement(lhs = Ft_Cost_Model_L_PSH_I24), StandardStatement(lhs = Ft_Cost_Model_L_PSH_I20), TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[4, "Q"])]
# =IF(C47="Underground",IF(G23="Yes",IF(O23,O23,200),0),0)*$C$106
Ft_Cost_Model_L_PSH_I23 = xl_mul((xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Underground") ? (xl_eq(inputs.Draft_Tube_Tunnels_Cost_Model_L_PSH_G23, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O23) ? inputs.s_Cost_Model_L_PSH_O23 : 200.0) : 0.0) : 0.0), No_Units__Cost_Model_L_PSH_C106) # Cost Model L-PSH I23
@assert xl_compare(Ft_Cost_Model_L_PSH_I23, 800) # "Cost Model L-PSH!I23"


# Level 6
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J22), StandardStatement(lhs = s_Cost_Model_L_PSH_J23)]
# "Cost Curves L-PSH!B131":"Cost Curves L-PSH!B132"
@. tab_Cost_Curves_L_PSH_B131_E132[!, "B"] = 47.162 * ((tab_Cost_Model_L_PSH_C108_C111[3:4, "C"]) ^ (1.6615))
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J22), StandardStatement(lhs = s_Cost_Model_L_PSH_J23)]
# "Cost Curves L-PSH!C131":"Cost Curves L-PSH!C132"
@. tab_Cost_Curves_L_PSH_B131_E132[!, "C"] = 43.711 * ((tab_Cost_Model_L_PSH_C108_C111[3:4, "C"]) ^ (1.7314))
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J22), StandardStatement(lhs = s_Cost_Model_L_PSH_J23)]
# "Cost Curves L-PSH!D131":"Cost Curves L-PSH!D132"
@. tab_Cost_Curves_L_PSH_B131_E132[!, "D"] = 45.454 * ((tab_Cost_Model_L_PSH_C108_C111[3:4, "C"]) ^ (1.8126))
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J22), StandardStatement(lhs = s_Cost_Model_L_PSH_J23)]
# "Cost Curves L-PSH!E131":"Cost Curves L-PSH!E132"
@. tab_Cost_Curves_L_PSH_B131_E132[!, "E"] = 60.182 * ((tab_Cost_Model_L_PSH_C108_C111[3:4, "C"]) ^ (1.7959))
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =3.9286*('Cost Model L-PSH'!C94)^2 + 10.071*('Cost Model L-PSH'!C94) + 481.43
s_Cost_Curves_L_PSH_V95 = 3.9286 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) + 10.071 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 481.43 # Cost Curves L-PSH V95
@assert xl_compare(s_Cost_Curves_L_PSH_V95, 3266.8581144914433) # "Cost Curves L-PSH!V95"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =4.3214*('Cost Model L-PSH'!C94)^2 - 5.6071*('Cost Model L-PSH'!C94) + 815
s_Cost_Curves_L_PSH_W95 = 4.3214 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) - 5.6071 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 815.0 # Cost Curves L-PSH W95
@assert xl_compare(s_Cost_Curves_L_PSH_W95, 3455.522413499425) # "Cost Curves L-PSH!W95"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =5.5536*('Cost Model L-PSH'!C94)^2 - 42.339*'Cost Model L-PSH'!C94 + 1165.4
s_Cost_Curves_L_PSH_X95 = 5.5536 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) - 42.339 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 1165.4 # Cost Curves L-PSH X95
@assert xl_compare(s_Cost_Curves_L_PSH_X95, 3667.287932744655) # "Cost Curves L-PSH!X95"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =7.6786*('Cost Model L-PSH'!C94)^2 - 103.54*'Cost Model L-PSH'!C94 + 1690.7
s_Cost_Curves_L_PSH_Y95 = 7.6786 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) - 103.54 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 1690.7 # Cost Curves L-PSH Y95
@assert xl_compare(s_Cost_Curves_L_PSH_Y95, 4007.944225216699) # "Cost Curves L-PSH!Y95"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =6.6786*('Cost Model L-PSH'!C94)^2 - 17.393*'Cost Model L-PSH'!C94 + 915
s_Cost_Curves_L_PSH_Z95 = 6.6786 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) - 17.393 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 915.0 # Cost Curves L-PSH Z95
@assert xl_compare(s_Cost_Curves_L_PSH_Z95, 4774.382357438666) # "Cost Curves L-PSH!Z95"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =6.25*('Cost Model L-PSH'!C94)^2 + 7.8929*'Cost Model L-PSH'!C94 + 819.29
s_Cost_Curves_L_PSH_AA95 = 6.25 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) + 7.8929 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 819.29 # Cost Curves L-PSH AA95
@assert xl_compare(s_Cost_Curves_L_PSH_AA95, 5044.335403452244) # "Cost Curves L-PSH!AA95"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =6.7143*('Cost Model L-PSH'!C94)^2 + 13.857*'Cost Model L-PSH'!C94 + 732.86
s_Cost_Curves_L_PSH_AB95 = 6.7143 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) + 13.857 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 732.86 # Cost Curves L-PSH AB95
@assert xl_compare(s_Cost_Curves_L_PSH_AB95, 5408.243428419412) # "Cost Curves L-PSH!AB95"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), StandardStatement(lhs = s_Cost_Model_L_PSH_J24)]
# =11.107*('Cost Model L-PSH'!C94)^2 - 120.11*'Cost Model L-PSH'!C94 + 1777.9
s_Cost_Curves_L_PSH_AC95 = 11.107 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) - 120.11 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 1777.9 # Cost Curves L-PSH AC95
@assert xl_compare(s_Cost_Curves_L_PSH_AC95, 5882.411391219839) # "Cost Curves L-PSH!AC95"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J21)]
# =186.57*'Cost Model L-PSH'!C94 + 27.143
_dollar_per_lf_Cost_Curves_L_PSH_AI90 = 186.57 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 27.143 # Cost Curves L-PSH AI90
@assert xl_compare(_dollar_per_lf_Cost_Curves_L_PSH_AI90, 4761.613089968174) # "Cost Curves L-PSH!AI90"
# Used in 6 places: [StandardStatement(lhs = kW_Cost_Model_L_PSH_I29), StandardStatement(lhs = kW_Cost_Model_L_PSH_I30), FunctionStatement(lhs = s_Cost_Model_L_PSH_J29), FunctionStatement(lhs = s_Cost_Model_L_PSH_J28), ..., FunctionStatement(lhs = s_Cost_Model_L_PSH_J30)]
tab_Cost_Model_L_PSH_C97_C105[8, "C"] = Mean_Gen_Discharge__Cost_Model_L_PSH_C87 * tab_Cost_Model_L_PSH_C97_C105[5, "C"] * inputs.P_T_Efficiency__Cost_Model_L_PSH_C71 * inputs.acceleration_of_gravity_metric_Cost_Model_L_PSH_C30 * inputs.density_of_water_metric_Cost_Model_L_PSH_C31 / inputs.MW_to_W_Cost_Model_L_PSH_C29 * ((1.0 / inputs.Meter_to_Feet_Cost_Model_L_PSH_C24) ^ (4.0)) # Cost Model L-PSH C104 Row: 8
@assert xl_compare(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 1005.030887807207) # "Cost Model L-PSH!C104"
# Used in 3 places: [StandardStatement(lhs = kW_Cost_Model_L_PSH_I29), StandardStatement(lhs = s_Cost_Curves_L_PSH_D374), StandardStatement(lhs = s_Cost_Model_L_PSH_Q10)]
# =IF(G10="Yes",IF(O10,O10,C105*1000),0)
kW_Cost_Model_L_PSH_I10 = (xl_eq(inputs.Powerplant_Structure_Cost_Model_L_PSH_G10, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O10) ? inputs.s_Cost_Model_L_PSH_O10 : tab_Cost_Model_L_PSH_C97_C105[9, "C"] * 1000.0) : 0.0) # Cost Model L-PSH I10
@assert xl_compare(kW_Cost_Model_L_PSH_I10, 1.2833248207034676e6) # "Cost Model L-PSH!I10"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J20), TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[1, "Q"])]
# =IF(G20="Yes",IF(O20,O20,(C15-C89-I22/C106-I23/C106)*0.5),0)*$C$93
Ft_Cost_Model_L_PSH_I20 = xl_mul((xl_eq(inputs.Upper_Low__and__High_Pressure_Tunnels_Cost_Model_L_PSH_G20, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O20) ? inputs.s_Cost_Model_L_PSH_O20 : xl_mul(xl_sub(xl_sub(inputs.Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15 - Mean_Gross_Head__Cost_Model_L_PSH_C89, xl_div(Ft_Cost_Model_L_PSH_I22, No_Units__Cost_Model_L_PSH_C106)), Ft_Cost_Model_L_PSH_I23 / No_Units__Cost_Model_L_PSH_C106), 0.5)) : 0.0), No_Tunnels__Cost_Model_L_PSH_C93) # Cost Model L-PSH I20
@assert xl_compare(Ft_Cost_Model_L_PSH_I20, 6268.25) # "Cost Model L-PSH!I20"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J24), TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[5, "Q"])]
# =IF(G24="Yes",IF(O24,O24,(C15-C89-I22/C106-I23/C106)*0.5),0)*$C$93
Ft_Cost_Model_L_PSH_I24 = xl_mul((xl_eq(inputs.Tailrace_Tunnels_Cost_Model_L_PSH_G24, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O24) ? inputs.s_Cost_Model_L_PSH_O24 : xl_mul(xl_sub(xl_sub(inputs.Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15 - Mean_Gross_Head__Cost_Model_L_PSH_C89, xl_div(Ft_Cost_Model_L_PSH_I22, No_Units__Cost_Model_L_PSH_C106)), Ft_Cost_Model_L_PSH_I23 / No_Units__Cost_Model_L_PSH_C106), 0.5)) : 0.0), No_Tunnels__Cost_Model_L_PSH_C93) # Cost Model L-PSH I24
@assert xl_compare(Ft_Cost_Model_L_PSH_I24, 6268.25) # "Cost Model L-PSH!I24"


# Level 7
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J34)]
# =2489.4*('Cost Model L-PSH'!C39)^0.118
s_Cost_Curves_L_PSH_D242 = 2489.4 * ((inputs.Access_Tunnel_Length_Cost_Model_L_PSH_C39) ^ (0.118)) # Cost Curves L-PSH D242
@assert xl_compare(s_Cost_Curves_L_PSH_D242, 2555.8188374587603) # "Cost Curves L-PSH!D242"
# Used in 1 places: [FunctionStatement(lhs = s_Cost_Model_L_PSH_J29)]
# ='Cost Model L-PSH'!I10/1000
s_Cost_Curves_L_PSH_D374 = kW_Cost_Model_L_PSH_I10 / 1000.0 # Cost Curves L-PSH D374
@assert xl_compare(s_Cost_Curves_L_PSH_D374, 1283.3248207034676) # "Cost Curves L-PSH!D374"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J14), StandardStatement(lhs = s_Cost_Model_L_PSH_J16)]
# =0.0154*('Cost Model L-PSH'!C94)^2 - 0.1968*'Cost Model L-PSH'!C94 + 1.85
horizontal_Cost_Curves_L_PSH_P88 = 0.0154 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) - 0.1968 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 1.85 # Cost Curves L-PSH P88
@assert xl_compare(horizontal_Cost_Curves_L_PSH_P88, 6.772919077968261) # "Cost Curves L-PSH!P88"
# Used in 2 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J14), StandardStatement(lhs = s_Cost_Model_L_PSH_J16)]
# =0.0046*('Cost Model L-PSH'!C94)^2 - 0.0819*'Cost Model L-PSH'!C94 + 0.5993
vertical_Cost_Curves_L_PSH_P91 = 0.0046 * ((Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94) ^ (2.0)) - 0.0819 * Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94 + 0.5993 # Cost Curves L-PSH P91
@assert xl_compare(vertical_Cost_Curves_L_PSH_P91, 1.4831926075256057) # "Cost Curves L-PSH!P91"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[1, "N"])]
# =IF(C34="Average",IF(I20*C25<=0.5,'Cost Curves L-PSH'!V95,IF(AND(I20*C25>0.5,I20*C25<=1),'Cost Curves L-PSH'!W95,IF(AND(I20*C25>1,I20*C25<=2),'Cost Curves L-PSH'!X95,IF(I20*C25>2,'Cost Curves L-PSH'!Y95)))),IF(C34="Poor",IF(I20*C25<=0.5,'Cost Curves L-PSH'!Z95,IF(AND(I20*C25>0.5,I20*C25<=1),'Cost Curves L-PSH'!AA95,IF(AND(I20*C25>1,I20*C25<=2),'Cost Curves L-PSH'!AB95,IF(I20*C25>2,'Cost Curves L-PSH'!AC95))))))
s_Cost_Model_L_PSH_J20 = (xl_eq(inputs.Tunneling_Condition_Cost_Model_L_PSH_C34, "Average") ? (xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5) ? s_Cost_Curves_L_PSH_V95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0)]) ? s_Cost_Curves_L_PSH_W95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0)]) ? s_Cost_Curves_L_PSH_X95 : (xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0) ? s_Cost_Curves_L_PSH_Y95 : missing)))) : (xl_eq(inputs.Tunneling_Condition_Cost_Model_L_PSH_C34, "Poor") ? (xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5) ? s_Cost_Curves_L_PSH_Z95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0)]) ? s_Cost_Curves_L_PSH_AA95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0)]) ? s_Cost_Curves_L_PSH_AB95 : (xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I20, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0) ? s_Cost_Curves_L_PSH_AC95 : missing)))) : missing)) # Cost Model L-PSH J20
@assert xl_compare(s_Cost_Model_L_PSH_J20, 3667.287932744655) # "Cost Model L-PSH!J20"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[2, "N"])]
# ='Cost Curves L-PSH'!AI90
s_Cost_Model_L_PSH_J21 = _dollar_per_lf_Cost_Curves_L_PSH_AI90 # Cost Model L-PSH J21
@assert xl_compare(s_Cost_Model_L_PSH_J21, 4761.613089968174) # "Cost Model L-PSH!J21"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[3, "N"])]
# =IF(C14<=500,'Cost Curves L-PSH'!B131,IF(AND(C14>500,C14<=1000),'Cost Curves L-PSH'!C131,IF(AND(C14>1000,C14<=1500),'Cost Curves L-PSH'!D131,IF(C14>1500,'Cost Curves L-PSH'!E131))))
s_Cost_Model_L_PSH_J22 = (xl_leq(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14, 500.0) ? tab_Cost_Curves_L_PSH_B131_E132[1, "B"] : (all([xl_gt(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14, 500.0), xl_leq(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14, 1000.0)]) ? tab_Cost_Curves_L_PSH_B131_E132[1, "C"] : (all([xl_gt(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14, 1000.0), xl_leq(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14, 1500.0)]) ? tab_Cost_Curves_L_PSH_B131_E132[1, "D"] : (xl_gt(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14, 1500.0) ? tab_Cost_Curves_L_PSH_B131_E132[1, "E"] : missing)))) # Cost Model L-PSH J22
@assert xl_compare(s_Cost_Model_L_PSH_J22, 4825.815230544194) # "Cost Model L-PSH!J22"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[4, "N"])]
# =IF(C47="Underground",IF(C14-C88<=500,'Cost Curves L-PSH'!B132,IF(AND(C14-C88>500,C14-C88<=1000),'Cost Curves L-PSH'!C132,IF(AND(C14-C88>1000,C14-C88<=1500),'Cost Curves L-PSH'!D132,IF(C14-C88>1500,'Cost Curves L-PSH'!E132)))),0)
s_Cost_Model_L_PSH_J23 = (xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Underground") ? (xl_leq(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - Min_Gross_Head__Cost_Model_L_PSH_C88, 500.0) ? tab_Cost_Curves_L_PSH_B131_E132[2, "B"] : (all([xl_gt(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - Min_Gross_Head__Cost_Model_L_PSH_C88, 500.0), xl_leq(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - Min_Gross_Head__Cost_Model_L_PSH_C88, 1000.0)]) ? tab_Cost_Curves_L_PSH_B131_E132[2, "C"] : (all([xl_gt(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - Min_Gross_Head__Cost_Model_L_PSH_C88, 1000.0), xl_leq(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - Min_Gross_Head__Cost_Model_L_PSH_C88, 1500.0)]) ? tab_Cost_Curves_L_PSH_B131_E132[2, "D"] : (xl_gt(inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14 - Min_Gross_Head__Cost_Model_L_PSH_C88, 1500.0) ? tab_Cost_Curves_L_PSH_B131_E132[2, "E"] : missing)))) : 0.0) # Cost Model L-PSH J23
@assert xl_compare(s_Cost_Model_L_PSH_J23, 6600.057681332245) # "Cost Model L-PSH!J23"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[5, "N"])]
# =IF(C34="Average",IF(I24*C25<=0.5,'Cost Curves L-PSH'!V95,IF(AND(I24*C25>0.5,I24*C25<=1),'Cost Curves L-PSH'!W95,IF(AND(I24*C25>1,I24*C25<=2),'Cost Curves L-PSH'!X95,IF(I24*C25>2,'Cost Curves L-PSH'!Y95)))),IF(C34="Poor",IF(I24*C25<=0.5,'Cost Curves L-PSH'!Z95,IF(AND(I24*C25>0.5,I24*C25<=1),'Cost Curves L-PSH'!AA95,IF(AND(I24*C25>1,I24*C25<=2),'Cost Curves L-PSH'!AB95,IF(I24*C25>2,'Cost Curves L-PSH'!AC95))))))
s_Cost_Model_L_PSH_J24 = (xl_eq(inputs.Tunneling_Condition_Cost_Model_L_PSH_C34, "Average") ? (xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5) ? s_Cost_Curves_L_PSH_V95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0)]) ? s_Cost_Curves_L_PSH_W95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0)]) ? s_Cost_Curves_L_PSH_X95 : (xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0) ? s_Cost_Curves_L_PSH_Y95 : missing)))) : (xl_eq(inputs.Tunneling_Condition_Cost_Model_L_PSH_C34, "Poor") ? (xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5) ? s_Cost_Curves_L_PSH_Z95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 0.5), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0)]) ? s_Cost_Curves_L_PSH_AA95 : (all([xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 1.0), xl_leq(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0)]) ? s_Cost_Curves_L_PSH_AB95 : (xl_gt(xl_mul(Ft_Cost_Model_L_PSH_I24, inputs.Feet_to_Miles_Cost_Model_L_PSH_C25), 2.0) ? s_Cost_Curves_L_PSH_AC95 : missing)))) : missing)) # Cost Model L-PSH J24
@assert xl_compare(s_Cost_Model_L_PSH_J24, 3667.287932744655) # "Cost Model L-PSH!J24"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[6, "N"])]
s_Cost_Model_L_PSH_J25 = calculate_s_Cost_Model_L_PSH_J25(inputs.Penstock_Cost_Model_L_PSH_C48, tab_Cost_Model_L_PSH_C97_C105, inputs.Nominal_Max_Head_Cost_Model_L_PSH_C14, tab_Cost_Model_L_PSH_C108_C111)
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[1, "N"])]
# =INDEX('Cost Curves L-PSH'!W241:X243,MATCH('Cost Model L-PSH'!C35,'Cost Curves L-PSH'!V241:V243,0),MATCH('Cost Model L-PSH'!C36,'Cost Curves L-PSH'!W240:X240,0))
s_Cost_Model_L_PSH_J33 = xl_index(tab_Cost_Curves_L_PSH_W241_X243[!, Between("W", "X")], xl_match(inputs.Access_Road_Terrain_Cost_Model_L_PSH_C35, [inputs.s_Cost_Curves_L_PSH_V241, inputs.s_Cost_Curves_L_PSH_V242, inputs.s_Cost_Curves_L_PSH_V243], 0.0), xl_match(inputs.Access_Road_Type_Cost_Model_L_PSH_C36, [inputs.Terrain_Cost_Curves_L_PSH_W240, inputs.New_Cost_Curves_L_PSH_X240], 0.0)) # Cost Model L-PSH J33
@assert xl_compare(s_Cost_Model_L_PSH_J33, 189000.0) # "Cost Model L-PSH!J33"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[1, "N"])]
tab_Cost_Model_L_PSH_K20_N25[1, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K20 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[1, "K"], 1.0) # "Cost Model L-PSH!K20"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[2, "N"])]
tab_Cost_Model_L_PSH_K20_N25[2, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K21 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[2, "K"], 1.0) # "Cost Model L-PSH!K21"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[3, "N"])]
tab_Cost_Model_L_PSH_K20_N25[3, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K22 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[3, "K"], 1.0) # "Cost Model L-PSH!K22"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[4, "N"])]
tab_Cost_Model_L_PSH_K20_N25[4, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K23 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[4, "K"], 1.0) # "Cost Model L-PSH!K23"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[5, "N"])]
tab_Cost_Model_L_PSH_K20_N25[5, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K24 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[5, "K"], 1.0) # "Cost Model L-PSH!K24"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[6, "N"])]
tab_Cost_Model_L_PSH_K20_N25[6, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K25 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[6, "K"], 1.0) # "Cost Model L-PSH!K25"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[1, "N"])]
tab_Cost_Model_L_PSH_K33_N34[1, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K33 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[1, "K"], 1.0) # "Cost Model L-PSH!K33"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[1, "N"])]
tab_Cost_Model_L_PSH_K20_N25[1, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L20 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[1, "L"], 2.4063214260550976) # "Cost Model L-PSH!L20"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[2, "N"])]
tab_Cost_Model_L_PSH_K20_N25[2, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L21 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[2, "L"], 2.4063214260550976) # "Cost Model L-PSH!L21"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[3, "N"])]
tab_Cost_Model_L_PSH_K20_N25[3, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L22 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[3, "L"], 2.4063214260550976) # "Cost Model L-PSH!L22"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[4, "N"])]
tab_Cost_Model_L_PSH_K20_N25[4, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L23 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[4, "L"], 2.4063214260550976) # "Cost Model L-PSH!L23"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[5, "N"])]
tab_Cost_Model_L_PSH_K20_N25[5, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L24 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[5, "L"], 2.4063214260550976) # "Cost Model L-PSH!L24"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[6, "N"])]
tab_Cost_Model_L_PSH_K20_N25[6, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L25 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[6, "L"], 2.4063214260550976) # "Cost Model L-PSH!L25"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[1, "N"])]
tab_Cost_Model_L_PSH_K33_N34[1, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L33 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[1, "L"], 2.4063214260550976) # "Cost Model L-PSH!L33"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[1, "N"])]
tab_Cost_Model_L_PSH_K20_N25[1, "M"] = inputs.Concrete_lined_tunnels_eg_tailrace__and__power_tun # Cost Model L-PSH M20 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[1, "M"], 1.6) # "Cost Model L-PSH!M20"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[2, "N"])]
tab_Cost_Model_L_PSH_K20_N25[2, "M"] = inputs.Vertical_shaft # Cost Model L-PSH M21 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[2, "M"], 1.8) # "Cost Model L-PSH!M21"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[3, "N"])]
tab_Cost_Model_L_PSH_K20_N25[3, "M"] = inputs.Steel_Lined_tunnels_underground_penstocks # Cost Model L-PSH M22 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[3, "M"], 1.9) # "Cost Model L-PSH!M22"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[4, "N"])]
tab_Cost_Model_L_PSH_K20_N25[4, "M"] = inputs.Draft_tubes # Cost Model L-PSH M23 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[4, "M"], 1.9) # "Cost Model L-PSH!M23"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[5, "N"])]
tab_Cost_Model_L_PSH_K20_N25[5, "M"] = inputs.Concrete_lined_tunnels_eg_tailrace__and__power_tun # Cost Model L-PSH M24 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[5, "M"], 1.6) # "Cost Model L-PSH!M24"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K20_N25[6, "N"])]
tab_Cost_Model_L_PSH_K20_N25[6, "M"] = inputs.Surface_penstocks # Cost Model L-PSH M25 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[6, "M"], 1.3) # "Cost Model L-PSH!M25"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[1, "N"])]
tab_Cost_Model_L_PSH_K33_N34[1, "M"] = inputs.Roads # Cost Model L-PSH M33 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[1, "M"], 1.0) # "Cost Model L-PSH!M33"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[1, "G"] = xl_average(tab_Land_Value_Data_B3_F3[!, Between("B", "F")]) # Land Value Data G3 Row: 1
@assert xl_compare(tab_Land_Value_Data_G3_G51[1, "G"], 9380) # "Land Value Data!G3"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[2, "G"] = xl_average(tab_Land_Value_Data_B4_F4[!, Between("B", "F")]) # Land Value Data G4 Row: 2
@assert xl_compare(tab_Land_Value_Data_G3_G51[2, "G"], 8256.666666666666) # "Land Value Data!G4"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[3, "G"] = xl_average(tab_Land_Value_Data_B5_F5[!, Between("B", "F")]) # Land Value Data G5 Row: 3
@assert xl_compare(tab_Land_Value_Data_G3_G51[3, "G"], 6080) # "Land Value Data!G5"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[4, "G"] = xl_average(tab_Land_Value_Data_B6_F6[!, Between("B", "F")]) # Land Value Data G6 Row: 4
@assert xl_compare(tab_Land_Value_Data_G3_G51[4, "G"], 8315) # "Land Value Data!G6"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[5, "G"] = xl_average(tab_Land_Value_Data_B7_F7[!, Between("B", "F")]) # Land Value Data G7 Row: 5
@assert xl_compare(tab_Land_Value_Data_G3_G51[5, "G"], 9780) # "Land Value Data!G7"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[6, "G"] = xl_average(tab_Land_Value_Data_B8_F8[!, Between("B", "F")]) # Land Value Data G8 Row: 6
@assert xl_compare(tab_Land_Value_Data_G3_G51[6, "G"], 6896.666666666667) # "Land Value Data!G8"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[7, "G"] = xl_average(tab_Land_Value_Data_B9_F9[!, Between("B", "F")]) # Land Value Data G9 Row: 7
@assert xl_compare(tab_Land_Value_Data_G3_G51[7, "G"], 14200) # "Land Value Data!G9"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[8, "G"] = xl_average(tab_Land_Value_Data_B10_F10[!, Between("B", "F")]) # Land Value Data G10 Row: 8
@assert xl_compare(tab_Land_Value_Data_G3_G51[8, "G"], 2586.6666666666665) # "Land Value Data!G10"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[9, "G"] = xl_average(tab_Land_Value_Data_B11_F11[!, Between("B", "F")]) # Land Value Data G11 Row: 9
@assert xl_compare(tab_Land_Value_Data_G3_G51[9, "G"], 5946.666666666667) # "Land Value Data!G11"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[10, "G"] = xl_average(tab_Land_Value_Data_B12_F12[!, Between("B", "F")]) # Land Value Data G12 Row: 10
@assert xl_compare(tab_Land_Value_Data_G3_G51[10, "G"], 10680) # "Land Value Data!G12"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[11, "G"] = xl_average(tab_Land_Value_Data_B13_F13[!, Between("B", "F")]) # Land Value Data G13 Row: 11
@assert xl_compare(tab_Land_Value_Data_G3_G51[11, "G"], 6513.333333333333) # "Land Value Data!G13"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[12, "G"] = xl_average(tab_Land_Value_Data_B14_F14[!, Between("B", "F")]) # Land Value Data G14 Row: 12
@assert xl_compare(tab_Land_Value_Data_G3_G51[12, "G"], 4246.666666666667) # "Land Value Data!G14"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[13, "G"] = xl_average(tab_Land_Value_Data_B15_F15[!, Between("B", "F")]) # Land Value Data G15 Row: 13
@assert xl_compare(tab_Land_Value_Data_G3_G51[13, "G"], 4113.333333333333) # "Land Value Data!G15"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[14, "G"] = xl_average(tab_Land_Value_Data_B16_F16[!, Between("B", "F")]) # Land Value Data G16 Row: 14
@assert xl_compare(tab_Land_Value_Data_G3_G51[14, "G"], 4330) # "Land Value Data!G16"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[15, "G"] = xl_average(tab_Land_Value_Data_B17_F17[!, Between("B", "F")]) # Land Value Data G17 Row: 15
@assert xl_compare(tab_Land_Value_Data_G3_G51[15, "G"], 6400) # "Land Value Data!G17"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[16, "G"] = xl_average(tab_Land_Value_Data_B18_F18[!, Between("B", "F")]) # Land Value Data G18 Row: 16
@assert xl_compare(tab_Land_Value_Data_G3_G51[16, "G"], 5463.333333333333) # "Land Value Data!G18"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[17, "G"] = xl_average(tab_Land_Value_Data_B19_F19[!, Between("B", "F")]) # Land Value Data G19 Row: 17
@assert xl_compare(tab_Land_Value_Data_G3_G51[17, "G"], 6190) # "Land Value Data!G19"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[18, "G"] = xl_average(tab_Land_Value_Data_B20_F20[!, Between("B", "F")]) # Land Value Data G20 Row: 18
@assert xl_compare(tab_Land_Value_Data_G3_G51[18, "G"], 3634) # "Land Value Data!G20"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[19, "G"] = xl_average(tab_Land_Value_Data_B21_F21[!, Between("B", "F")]) # Land Value Data G21 Row: 19
@assert xl_compare(tab_Land_Value_Data_G3_G51[19, "G"], 5613.333333333333) # "Land Value Data!G21"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[20, "G"] = xl_average(tab_Land_Value_Data_B22_F22[!, Between("B", "F")]) # Land Value Data G22 Row: 20
@assert xl_compare(tab_Land_Value_Data_G3_G51[20, "G"], 2384) # "Land Value Data!G22"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[21, "G"] = xl_average(tab_Land_Value_Data_B23_F23[!, Between("B", "F")]) # Land Value Data G23 Row: 21
@assert xl_compare(tab_Land_Value_Data_G3_G51[21, "G"], 3932) # "Land Value Data!G23"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[22, "G"] = xl_average(tab_Land_Value_Data_B24_F24[!, Between("B", "F")]) # Land Value Data G24 Row: 22
@assert xl_compare(tab_Land_Value_Data_G3_G51[22, "G"], 1573.3333333333333) # "Land Value Data!G24"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[23, "G"] = xl_average(tab_Land_Value_Data_B25_F25[!, Between("B", "F")]) # Land Value Data G25 Row: 23
@assert xl_compare(tab_Land_Value_Data_G3_G51[23, "G"], 2455) # "Land Value Data!G25"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[24, "G"] = xl_average(tab_Land_Value_Data_B26_F26[!, Between("B", "F")]) # Land Value Data G26 Row: 24
@assert xl_compare(tab_Land_Value_Data_G3_G51[24, "G"], 3863.3333333333335) # "Land Value Data!G26"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[25, "G"] = xl_average(tab_Land_Value_Data_B27_F27[!, Between("B", "F")]) # Land Value Data G27 Row: 25
@assert xl_compare(tab_Land_Value_Data_G3_G51[25, "G"], 4630) # "Land Value Data!G27"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[26, "G"] = xl_average(tab_Land_Value_Data_B28_F28[!, Between("B", "F")]) # Land Value Data G28 Row: 26
@assert xl_compare(tab_Land_Value_Data_G3_G51[26, "G"], 4130) # "Land Value Data!G28"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[27, "G"] = xl_average(tab_Land_Value_Data_B29_F29[!, Between("B", "F")]) # Land Value Data G29 Row: 27
@assert xl_compare(tab_Land_Value_Data_G3_G51[27, "G"], 4516.666666666667) # "Land Value Data!G29"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[28, "G"] = xl_average(tab_Land_Value_Data_B30_F30[!, Between("B", "F")]) # Land Value Data G30 Row: 28
@assert xl_compare(tab_Land_Value_Data_G3_G51[28, "G"], 2766.6666666666665) # "Land Value Data!G30"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[29, "G"] = xl_average(tab_Land_Value_Data_B31_F31[!, Between("B", "F")]) # Land Value Data G31 Row: 29
@assert xl_compare(tab_Land_Value_Data_G3_G51[29, "G"], 3133.3333333333335) # "Land Value Data!G31"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[30, "G"] = xl_average(tab_Land_Value_Data_B32_F32[!, Between("B", "F")]) # Land Value Data G32 Row: 30
@assert xl_compare(tab_Land_Value_Data_G3_G51[30, "G"], 6704) # "Land Value Data!G32"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[31, "G"] = xl_average(tab_Land_Value_Data_B33_F33[!, Between("B", "F")]) # Land Value Data G33 Row: 31
@assert xl_compare(tab_Land_Value_Data_G3_G51[31, "G"], 3738) # "Land Value Data!G33"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[32, "G"] = xl_average(tab_Land_Value_Data_B34_F34[!, Between("B", "F")]) # Land Value Data G34 Row: 32
@assert xl_compare(tab_Land_Value_Data_G3_G51[32, "G"], 3283.3333333333335) # "Land Value Data!G34"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[33, "G"] = xl_average(tab_Land_Value_Data_B35_F35[!, Between("B", "F")]) # Land Value Data G35 Row: 33
@assert xl_compare(tab_Land_Value_Data_G3_G51[33, "G"], 2914) # "Land Value Data!G35"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[34, "G"] = xl_average(tab_Land_Value_Data_B36_F36[!, Between("B", "F")]) # Land Value Data G36 Row: 34
@assert xl_compare(tab_Land_Value_Data_G3_G51[34, "G"], 3010) # "Land Value Data!G36"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[35, "G"] = xl_average(tab_Land_Value_Data_B37_F37[!, Between("B", "F")]) # Land Value Data G37 Row: 35
@assert xl_compare(tab_Land_Value_Data_G3_G51[35, "G"], 3004) # "Land Value Data!G37"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[36, "G"] = xl_average(tab_Land_Value_Data_B38_F38[!, Between("B", "F")]) # Land Value Data G38 Row: 36
@assert xl_compare(tab_Land_Value_Data_G3_G51[36, "G"], 1805) # "Land Value Data!G38"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[37, "G"] = xl_average(tab_Land_Value_Data_B39_F39[!, Between("B", "F")]) # Land Value Data G39 Row: 37
@assert xl_compare(tab_Land_Value_Data_G3_G51[37, "G"], 2192) # "Land Value Data!G39"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[38, "G"] = xl_average(tab_Land_Value_Data_B40_F40[!, Between("B", "F")]) # Land Value Data G40 Row: 38
@assert xl_compare(tab_Land_Value_Data_G3_G51[38, "G"], 6433.333333333333) # "Land Value Data!G40"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[39, "G"] = xl_average(tab_Land_Value_Data_B41_F41[!, Between("B", "F")]) # Land Value Data G41 Row: 39
@assert xl_compare(tab_Land_Value_Data_G3_G51[39, "G"], 2305) # "Land Value Data!G41"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[40, "G"] = xl_average(tab_Land_Value_Data_B42_F42[!, Between("B", "F")]) # Land Value Data G42 Row: 40
@assert xl_compare(tab_Land_Value_Data_G3_G51[40, "G"], 3638) # "Land Value Data!G42"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[41, "G"] = xl_average(tab_Land_Value_Data_B43_F43[!, Between("B", "F")]) # Land Value Data G43 Row: 41
@assert xl_compare(tab_Land_Value_Data_G3_G51[41, "G"], 1313) # "Land Value Data!G43"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[42, "G"] = xl_average(tab_Land_Value_Data_B44_F44[!, Between("B", "F")]) # Land Value Data G44 Row: 42
@assert xl_compare(tab_Land_Value_Data_G3_G51[42, "G"], 1010) # "Land Value Data!G44"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[43, "G"] = xl_average(tab_Land_Value_Data_B45_F45[!, Between("B", "F")]) # Land Value Data G45 Row: 43
@assert xl_compare(tab_Land_Value_Data_G3_G51[43, "G"], 1547) # "Land Value Data!G45"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[44, "G"] = xl_average(tab_Land_Value_Data_B46_F46[!, Between("B", "F")]) # Land Value Data G46 Row: 44
@assert xl_compare(tab_Land_Value_Data_G3_G51[44, "G"], 3276) # "Land Value Data!G46"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[45, "G"] = xl_average(tab_Land_Value_Data_B47_F47[!, Between("B", "F")]) # Land Value Data G47 Row: 45
@assert xl_compare(tab_Land_Value_Data_G3_G51[45, "G"], 1288) # "Land Value Data!G47"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[46, "G"] = xl_average(tab_Land_Value_Data_B48_F48[!, Between("B", "F")]) # Land Value Data G48 Row: 46
@assert xl_compare(tab_Land_Value_Data_G3_G51[46, "G"], 10012) # "Land Value Data!G48"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[47, "G"] = xl_average(tab_Land_Value_Data_B49_F49[!, Between("B", "F")]) # Land Value Data G49 Row: 47
@assert xl_compare(tab_Land_Value_Data_G3_G51[47, "G"], 3014) # "Land Value Data!G49"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[48, "G"] = xl_average(tab_Land_Value_Data_B50_F50[!, Between("B", "F")]) # Land Value Data G50 Row: 48
@assert xl_compare(tab_Land_Value_Data_G3_G51[48, "G"], 3092) # "Land Value Data!G50"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_J8)]
tab_Land_Value_Data_G3_G51[49, "G"] = xl_average(tab_Land_Value_Data_B51_F51[!, Between("B", "F")]) # Land Value Data G51 Row: 49
@assert xl_compare(tab_Land_Value_Data_G3_G51[49, "G"], 3093.3333333333335) # "Land Value Data!G51"


# Level 8
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_I13_I14[1, "I"])]
tab_Cost_Model_L_PSH_C80_C86[4, "C"] = (9.0e-5 * ((inputs.Avg_Upper_Dam_Height_Cost_Model_L_PSH_C16) ^ (2.0)) + 0.0039 * inputs.Avg_Upper_Dam_Height_Cost_Model_L_PSH_C16 + 0.0707) * 1000.0 * inputs.Upper_Dam_Crest_Length_Cost_Model_L_PSH_C17 # Cost Model L-PSH C83 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_C80_C86[4, "C"], 2385110) # "Cost Model L-PSH!C83"
# Used in 1 places: [StandardStatement(lhs = CY_Cost_Model_L_PSH_I17)]
tab_Cost_Model_L_PSH_C80_C86[5, "C"] = (9.0e-5 * ((inputs.Avg_Lower_Dam_Height_Cost_Model_L_PSH_C18) ^ (2.0)) + 0.0039 * inputs.Avg_Lower_Dam_Height_Cost_Model_L_PSH_C18 + 0.0707) * 1000.0 * inputs.Lower_Dam_Crest_Length_Cost_Model_L_PSH_C19 # Cost Model L-PSH C84 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_C80_C86[5, "C"], 691570) # "Cost Model L-PSH!C84"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[2, "Q"])]
# =IF(G21="Yes",IF(O21,O21,C89),0)*$C$93
Ft_Cost_Model_L_PSH_I21 = xl_mul((xl_eq(inputs.Vertical_Shafts_Cost_Model_L_PSH_G21, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O21) ? inputs.s_Cost_Model_L_PSH_O21 : Mean_Gross_Head__Cost_Model_L_PSH_C89) : 0.0), No_Tunnels__Cost_Model_L_PSH_C93) # Cost Model L-PSH I21
@assert xl_compare(Ft_Cost_Model_L_PSH_I21, 1326) # "Cost Model L-PSH!I21"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[6, "Q"])]
# =IF(G25="Yes",IF(O25,O25,IF(C48="Surface",C106*C89*0.25,0)),0)
Ft_Cost_Model_L_PSH_I25 = (xl_eq(inputs.Surface_Penstock_Cost_Model_L_PSH_G25, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O25) ? inputs.s_Cost_Model_L_PSH_O25 : (xl_eq(inputs.Penstock_Cost_Model_L_PSH_C48, "Surface") ? No_Units__Cost_Model_L_PSH_C106 * Mean_Gross_Head__Cost_Model_L_PSH_C89 * 0.25 : 0.0)) : 0.0) # Cost Model L-PSH I25
@assert xl_compare(Ft_Cost_Model_L_PSH_I25, 0) # "Cost Model L-PSH!I25"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q33_R35[1, "Q"])]
# =IF(G33="Yes",IF(O33,O33,C38),0)
Miles_Cost_Model_L_PSH_I33 = (xl_eq(inputs.Access_Roads_Cost_Model_L_PSH_G33, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O33) ? inputs.s_Cost_Model_L_PSH_O33 : inputs.Access_Road_Cost_Model_L_PSH_C38) : 0.0) # Cost Model L-PSH I33
@assert xl_compare(Miles_Cost_Model_L_PSH_I33, 3.61) # "Cost Model L-PSH!I33"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_N8)]
# =IFERROR(VLOOKUP(C9,'Land Value Data'!$A$3:$G$51,7,FALSE),'Land Value Data'!G51)
s_Cost_Model_L_PSH_J8 = xl_iferror(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, [[inputs.s_Land_Value_Data_A3, inputs.s_Land_Value_Data_A4, inputs.s_Land_Value_Data_A5, inputs.s_Land_Value_Data_A6, inputs.s_Land_Value_Data_A7, inputs.s_Land_Value_Data_A8, inputs.s_Land_Value_Data_A9, inputs.s_Land_Value_Data_A10, inputs.s_Land_Value_Data_A11, inputs.s_Land_Value_Data_A12, inputs.s_Land_Value_Data_A13, inputs.s_Land_Value_Data_A14, inputs.s_Land_Value_Data_A15, inputs.s_Land_Value_Data_A16, inputs.s_Land_Value_Data_A17, inputs.s_Land_Value_Data_A18, inputs.s_Land_Value_Data_A19, inputs.s_Land_Value_Data_A20, inputs.s_Land_Value_Data_A21, inputs.s_Land_Value_Data_A22, inputs.s_Land_Value_Data_A23, inputs.s_Land_Value_Data_A24, inputs.s_Land_Value_Data_A25, inputs.s_Land_Value_Data_A26, inputs.s_Land_Value_Data_A27, inputs.s_Land_Value_Data_A28, inputs.s_Land_Value_Data_A29, inputs.s_Land_Value_Data_A30, inputs.s_Land_Value_Data_A31, inputs.s_Land_Value_Data_A32, inputs.s_Land_Value_Data_A33, inputs.s_Land_Value_Data_A34, inputs.s_Land_Value_Data_A35, inputs.s_Land_Value_Data_A36, inputs.s_Land_Value_Data_A37, inputs.s_Land_Value_Data_A38, inputs.s_Land_Value_Data_A39, inputs.s_Land_Value_Data_A40, inputs.s_Land_Value_Data_A41, inputs.s_Land_Value_Data_A42, inputs.s_Land_Value_Data_A43, inputs.s_Land_Value_Data_A44, inputs.s_Land_Value_Data_A45, inputs.s_Land_Value_Data_A46, inputs.s_Land_Value_Data_A47, inputs.s_Land_Value_Data_A48, inputs.s_Land_Value_Data_A49, inputs.s_Land_Value_Data_A50, inputs.s_Land_Value_Data_A51] [tab_Land_Value_Data_B3_F3[1, "B"], tab_Land_Value_Data_B4_F4[1, "B"], tab_Land_Value_Data_B5_F5[1, "B"], tab_Land_Value_Data_B6_F6[1, "B"], tab_Land_Value_Data_B7_F7[1, "B"], tab_Land_Value_Data_B8_F8[1, "B"], tab_Land_Value_Data_B9_F9[1, "B"], tab_Land_Value_Data_B10_F10[1, "B"], tab_Land_Value_Data_B11_F11[1, "B"], tab_Land_Value_Data_B12_F12[1, "B"], tab_Land_Value_Data_B13_F13[1, "B"], tab_Land_Value_Data_B14_F14[1, "B"], tab_Land_Value_Data_B15_F15[1, "B"], tab_Land_Value_Data_B16_F16[1, "B"], tab_Land_Value_Data_B17_F17[1, "B"], tab_Land_Value_Data_B18_F18[1, "B"], tab_Land_Value_Data_B19_F19[1, "B"], tab_Land_Value_Data_B20_F20[1, "B"], tab_Land_Value_Data_B21_F21[1, "B"], tab_Land_Value_Data_B22_F22[1, "B"], tab_Land_Value_Data_B23_F23[1, "B"], tab_Land_Value_Data_B24_F24[1, "B"], tab_Land_Value_Data_B25_F25[1, "B"], tab_Land_Value_Data_B26_F26[1, "B"], tab_Land_Value_Data_B27_F27[1, "B"], tab_Land_Value_Data_B28_F28[1, "B"], tab_Land_Value_Data_B29_F29[1, "B"], tab_Land_Value_Data_B30_F30[1, "B"], tab_Land_Value_Data_B31_F31[1, "B"], tab_Land_Value_Data_B32_F32[1, "B"], tab_Land_Value_Data_B33_F33[1, "B"], tab_Land_Value_Data_B34_F34[1, "B"], tab_Land_Value_Data_B35_F35[1, "B"], tab_Land_Value_Data_B36_F36[1, "B"], tab_Land_Value_Data_B37_F37[1, "B"], tab_Land_Value_Data_B38_F38[1, "B"], tab_Land_Value_Data_B39_F39[1, "B"], tab_Land_Value_Data_B40_F40[1, "B"], tab_Land_Value_Data_B41_F41[1, "B"], tab_Land_Value_Data_B42_F42[1, "B"], tab_Land_Value_Data_B43_F43[1, "B"], tab_Land_Value_Data_B44_F44[1, "B"], tab_Land_Value_Data_B45_F45[1, "B"], tab_Land_Value_Data_B46_F46[1, "B"], tab_Land_Value_Data_B47_F47[1, "B"], tab_Land_Value_Data_B48_F48[1, "B"], tab_Land_Value_Data_B49_F49[1, "B"], tab_Land_Value_Data_B50_F50[1, "B"], tab_Land_Value_Data_B51_F51[1, "B"]] [tab_Land_Value_Data_B3_F3[1, "C"], tab_Land_Value_Data_B4_F4[1, "C"], tab_Land_Value_Data_B5_F5[1, "C"], tab_Land_Value_Data_B6_F6[1, "C"], tab_Land_Value_Data_B7_F7[1, "C"], tab_Land_Value_Data_B8_F8[1, "C"], tab_Land_Value_Data_B9_F9[1, "C"], tab_Land_Value_Data_B10_F10[1, "C"], tab_Land_Value_Data_B11_F11[1, "C"], tab_Land_Value_Data_B12_F12[1, "C"], tab_Land_Value_Data_B13_F13[1, "C"], tab_Land_Value_Data_B14_F14[1, "C"], tab_Land_Value_Data_B15_F15[1, "C"], tab_Land_Value_Data_B16_F16[1, "C"], tab_Land_Value_Data_B17_F17[1, "C"], tab_Land_Value_Data_B18_F18[1, "C"], tab_Land_Value_Data_B19_F19[1, "C"], tab_Land_Value_Data_B20_F20[1, "C"], tab_Land_Value_Data_B21_F21[1, "C"], tab_Land_Value_Data_B22_F22[1, "C"], tab_Land_Value_Data_B23_F23[1, "C"], tab_Land_Value_Data_B24_F24[1, "C"], tab_Land_Value_Data_B25_F25[1, "C"], tab_Land_Value_Data_B26_F26[1, "C"], tab_Land_Value_Data_B27_F27[1, "C"], tab_Land_Value_Data_B28_F28[1, "C"], tab_Land_Value_Data_B29_F29[1, "C"], tab_Land_Value_Data_B30_F30[1, "C"], tab_Land_Value_Data_B31_F31[1, "C"], tab_Land_Value_Data_B32_F32[1, "C"], tab_Land_Value_Data_B33_F33[1, "C"], tab_Land_Value_Data_B34_F34[1, "C"], tab_Land_Value_Data_B35_F35[1, "C"], tab_Land_Value_Data_B36_F36[1, "C"], tab_Land_Value_Data_B37_F37[1, "C"], tab_Land_Value_Data_B38_F38[1, "C"], tab_Land_Value_Data_B39_F39[1, "C"], tab_Land_Value_Data_B40_F40[1, "C"], tab_Land_Value_Data_B41_F41[1, "C"], tab_Land_Value_Data_B42_F42[1, "C"], tab_Land_Value_Data_B43_F43[1, "C"], tab_Land_Value_Data_B44_F44[1, "C"], tab_Land_Value_Data_B45_F45[1, "C"], tab_Land_Value_Data_B46_F46[1, "C"], tab_Land_Value_Data_B47_F47[1, "C"], tab_Land_Value_Data_B48_F48[1, "C"], tab_Land_Value_Data_B49_F49[1, "C"], tab_Land_Value_Data_B50_F50[1, "C"], tab_Land_Value_Data_B51_F51[1, "C"]] [tab_Land_Value_Data_B3_F3[1, "D"], tab_Land_Value_Data_B4_F4[1, "D"], tab_Land_Value_Data_B5_F5[1, "D"], tab_Land_Value_Data_B6_F6[1, "D"], tab_Land_Value_Data_B7_F7[1, "D"], tab_Land_Value_Data_B8_F8[1, "D"], tab_Land_Value_Data_B9_F9[1, "D"], tab_Land_Value_Data_B10_F10[1, "D"], tab_Land_Value_Data_B11_F11[1, "D"], tab_Land_Value_Data_B12_F12[1, "D"], tab_Land_Value_Data_B13_F13[1, "D"], tab_Land_Value_Data_B14_F14[1, "D"], tab_Land_Value_Data_B15_F15[1, "D"], tab_Land_Value_Data_B16_F16[1, "D"], tab_Land_Value_Data_B17_F17[1, "D"], tab_Land_Value_Data_B18_F18[1, "D"], tab_Land_Value_Data_B19_F19[1, "D"], tab_Land_Value_Data_B20_F20[1, "D"], tab_Land_Value_Data_B21_F21[1, "D"], tab_Land_Value_Data_B22_F22[1, "D"], tab_Land_Value_Data_B23_F23[1, "D"], tab_Land_Value_Data_B24_F24[1, "D"], tab_Land_Value_Data_B25_F25[1, "D"], tab_Land_Value_Data_B26_F26[1, "D"], tab_Land_Value_Data_B27_F27[1, "D"], tab_Land_Value_Data_B28_F28[1, "D"], tab_Land_Value_Data_B29_F29[1, "D"], tab_Land_Value_Data_B30_F30[1, "D"], tab_Land_Value_Data_B31_F31[1, "D"], tab_Land_Value_Data_B32_F32[1, "D"], tab_Land_Value_Data_B33_F33[1, "D"], tab_Land_Value_Data_B34_F34[1, "D"], tab_Land_Value_Data_B35_F35[1, "D"], tab_Land_Value_Data_B36_F36[1, "D"], tab_Land_Value_Data_B37_F37[1, "D"], tab_Land_Value_Data_B38_F38[1, "D"], tab_Land_Value_Data_B39_F39[1, "D"], tab_Land_Value_Data_B40_F40[1, "D"], tab_Land_Value_Data_B41_F41[1, "D"], tab_Land_Value_Data_B42_F42[1, "D"], tab_Land_Value_Data_B43_F43[1, "D"], tab_Land_Value_Data_B44_F44[1, "D"], tab_Land_Value_Data_B45_F45[1, "D"], tab_Land_Value_Data_B46_F46[1, "D"], tab_Land_Value_Data_B47_F47[1, "D"], tab_Land_Value_Data_B48_F48[1, "D"], tab_Land_Value_Data_B49_F49[1, "D"], tab_Land_Value_Data_B50_F50[1, "D"], tab_Land_Value_Data_B51_F51[1, "D"]] [tab_Land_Value_Data_B3_F3[1, "E"], tab_Land_Value_Data_B4_F4[1, "E"], tab_Land_Value_Data_B5_F5[1, "E"], tab_Land_Value_Data_B6_F6[1, "E"], tab_Land_Value_Data_B7_F7[1, "E"], tab_Land_Value_Data_B8_F8[1, "E"], tab_Land_Value_Data_B9_F9[1, "E"], tab_Land_Value_Data_B10_F10[1, "E"], tab_Land_Value_Data_B11_F11[1, "E"], tab_Land_Value_Data_B12_F12[1, "E"], tab_Land_Value_Data_B13_F13[1, "E"], tab_Land_Value_Data_B14_F14[1, "E"], tab_Land_Value_Data_B15_F15[1, "E"], tab_Land_Value_Data_B16_F16[1, "E"], tab_Land_Value_Data_B17_F17[1, "E"], tab_Land_Value_Data_B18_F18[1, "E"], tab_Land_Value_Data_B19_F19[1, "E"], tab_Land_Value_Data_B20_F20[1, "E"], tab_Land_Value_Data_B21_F21[1, "E"], tab_Land_Value_Data_B22_F22[1, "E"], tab_Land_Value_Data_B23_F23[1, "E"], tab_Land_Value_Data_B24_F24[1, "E"], tab_Land_Value_Data_B25_F25[1, "E"], tab_Land_Value_Data_B26_F26[1, "E"], tab_Land_Value_Data_B27_F27[1, "E"], tab_Land_Value_Data_B28_F28[1, "E"], tab_Land_Value_Data_B29_F29[1, "E"], tab_Land_Value_Data_B30_F30[1, "E"], tab_Land_Value_Data_B31_F31[1, "E"], tab_Land_Value_Data_B32_F32[1, "E"], tab_Land_Value_Data_B33_F33[1, "E"], tab_Land_Value_Data_B34_F34[1, "E"], tab_Land_Value_Data_B35_F35[1, "E"], tab_Land_Value_Data_B36_F36[1, "E"], tab_Land_Value_Data_B37_F37[1, "E"], tab_Land_Value_Data_B38_F38[1, "E"], tab_Land_Value_Data_B39_F39[1, "E"], tab_Land_Value_Data_B40_F40[1, "E"], tab_Land_Value_Data_B41_F41[1, "E"], tab_Land_Value_Data_B42_F42[1, "E"], tab_Land_Value_Data_B43_F43[1, "E"], tab_Land_Value_Data_B44_F44[1, "E"], tab_Land_Value_Data_B45_F45[1, "E"], tab_Land_Value_Data_B46_F46[1, "E"], tab_Land_Value_Data_B47_F47[1, "E"], tab_Land_Value_Data_B48_F48[1, "E"], tab_Land_Value_Data_B49_F49[1, "E"], tab_Land_Value_Data_B50_F50[1, "E"], tab_Land_Value_Data_B51_F51[1, "E"]] [tab_Land_Value_Data_B3_F3[1, "F"], tab_Land_Value_Data_B4_F4[1, "F"], tab_Land_Value_Data_B5_F5[1, "F"], tab_Land_Value_Data_B6_F6[1, "F"], tab_Land_Value_Data_B7_F7[1, "F"], tab_Land_Value_Data_B8_F8[1, "F"], tab_Land_Value_Data_B9_F9[1, "F"], tab_Land_Value_Data_B10_F10[1, "F"], tab_Land_Value_Data_B11_F11[1, "F"], tab_Land_Value_Data_B12_F12[1, "F"], tab_Land_Value_Data_B13_F13[1, "F"], tab_Land_Value_Data_B14_F14[1, "F"], tab_Land_Value_Data_B15_F15[1, "F"], tab_Land_Value_Data_B16_F16[1, "F"], tab_Land_Value_Data_B17_F17[1, "F"], tab_Land_Value_Data_B18_F18[1, "F"], tab_Land_Value_Data_B19_F19[1, "F"], tab_Land_Value_Data_B20_F20[1, "F"], tab_Land_Value_Data_B21_F21[1, "F"], tab_Land_Value_Data_B22_F22[1, "F"], tab_Land_Value_Data_B23_F23[1, "F"], tab_Land_Value_Data_B24_F24[1, "F"], tab_Land_Value_Data_B25_F25[1, "F"], tab_Land_Value_Data_B26_F26[1, "F"], tab_Land_Value_Data_B27_F27[1, "F"], tab_Land_Value_Data_B28_F28[1, "F"], tab_Land_Value_Data_B29_F29[1, "F"], tab_Land_Value_Data_B30_F30[1, "F"], tab_Land_Value_Data_B31_F31[1, "F"], tab_Land_Value_Data_B32_F32[1, "F"], tab_Land_Value_Data_B33_F33[1, "F"], tab_Land_Value_Data_B34_F34[1, "F"], tab_Land_Value_Data_B35_F35[1, "F"], tab_Land_Value_Data_B36_F36[1, "F"], tab_Land_Value_Data_B37_F37[1, "F"], tab_Land_Value_Data_B38_F38[1, "F"], tab_Land_Value_Data_B39_F39[1, "F"], tab_Land_Value_Data_B40_F40[1, "F"], tab_Land_Value_Data_B41_F41[1, "F"], tab_Land_Value_Data_B42_F42[1, "F"], tab_Land_Value_Data_B43_F43[1, "F"], tab_Land_Value_Data_B44_F44[1, "F"], tab_Land_Value_Data_B45_F45[1, "F"], tab_Land_Value_Data_B46_F46[1, "F"], tab_Land_Value_Data_B47_F47[1, "F"], tab_Land_Value_Data_B48_F48[1, "F"], tab_Land_Value_Data_B49_F49[1, "F"], tab_Land_Value_Data_B50_F50[1, "F"], tab_Land_Value_Data_B51_F51[1, "F"]] [tab_Land_Value_Data_G3_G51[1, "G"], tab_Land_Value_Data_G3_G51[2, "G"], tab_Land_Value_Data_G3_G51[3, "G"], tab_Land_Value_Data_G3_G51[4, "G"], tab_Land_Value_Data_G3_G51[5, "G"], tab_Land_Value_Data_G3_G51[6, "G"], tab_Land_Value_Data_G3_G51[7, "G"], tab_Land_Value_Data_G3_G51[8, "G"], tab_Land_Value_Data_G3_G51[9, "G"], tab_Land_Value_Data_G3_G51[10, "G"], tab_Land_Value_Data_G3_G51[11, "G"], tab_Land_Value_Data_G3_G51[12, "G"], tab_Land_Value_Data_G3_G51[13, "G"], tab_Land_Value_Data_G3_G51[14, "G"], tab_Land_Value_Data_G3_G51[15, "G"], tab_Land_Value_Data_G3_G51[16, "G"], tab_Land_Value_Data_G3_G51[17, "G"], tab_Land_Value_Data_G3_G51[18, "G"], tab_Land_Value_Data_G3_G51[19, "G"], tab_Land_Value_Data_G3_G51[20, "G"], tab_Land_Value_Data_G3_G51[21, "G"], tab_Land_Value_Data_G3_G51[22, "G"], tab_Land_Value_Data_G3_G51[23, "G"], tab_Land_Value_Data_G3_G51[24, "G"], tab_Land_Value_Data_G3_G51[25, "G"], tab_Land_Value_Data_G3_G51[26, "G"], tab_Land_Value_Data_G3_G51[27, "G"], tab_Land_Value_Data_G3_G51[28, "G"], tab_Land_Value_Data_G3_G51[29, "G"], tab_Land_Value_Data_G3_G51[30, "G"], tab_Land_Value_Data_G3_G51[31, "G"], tab_Land_Value_Data_G3_G51[32, "G"], tab_Land_Value_Data_G3_G51[33, "G"], tab_Land_Value_Data_G3_G51[34, "G"], tab_Land_Value_Data_G3_G51[35, "G"], tab_Land_Value_Data_G3_G51[36, "G"], tab_Land_Value_Data_G3_G51[37, "G"], tab_Land_Value_Data_G3_G51[38, "G"], tab_Land_Value_Data_G3_G51[39, "G"], tab_Land_Value_Data_G3_G51[40, "G"], tab_Land_Value_Data_G3_G51[41, "G"], tab_Land_Value_Data_G3_G51[42, "G"], tab_Land_Value_Data_G3_G51[43, "G"], tab_Land_Value_Data_G3_G51[44, "G"], tab_Land_Value_Data_G3_G51[45, "G"], tab_Land_Value_Data_G3_G51[46, "G"], tab_Land_Value_Data_G3_G51[47, "G"], tab_Land_Value_Data_G3_G51[48, "G"], tab_Land_Value_Data_G3_G51[49, "G"]]], 7.0, false), tab_Land_Value_Data_G3_G51[49, "G"]) # Cost Model L-PSH J8
@assert xl_compare(s_Cost_Model_L_PSH_J8, 3093.3333333333335) # "Cost Model L-PSH!J8"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[1, "N"])]
s_Cost_Model_L_PSH_J13 = calculate_s_Cost_Model_L_PSH_J13(inputs.Upper_Dam_Crest_Length_Cost_Model_L_PSH_C17, inputs.Avg_Upper_Dam_Height_Cost_Model_L_PSH_C16, tab_Cost_Curves_L_PSH_J95_J96)
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[2, "N"])]
# =IF(C44="Horizontal",'Cost Curves L-PSH'!P88,IF('Cost Model L-PSH'!C44="Vertical",'Cost Curves L-PSH'!P91))*1000000
s_Cost_Model_L_PSH_J14 = xl_mul((xl_eq(inputs.U_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_C44, "Horizontal") ? horizontal_Cost_Curves_L_PSH_P88 : (xl_eq(inputs.U_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_C44, "Vertical") ? vertical_Cost_Curves_L_PSH_P91 : missing)), 1.0e6) # Cost Model L-PSH J14
@assert xl_compare(s_Cost_Model_L_PSH_J14, 1.4831926075256057e6) # "Cost Model L-PSH!J14"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[1, "N"])]
# =IF(C45="Horizontal",'Cost Curves L-PSH'!P88,IF('Cost Model L-PSH'!C45="Vertical",'Cost Curves L-PSH'!P91))*1000000
s_Cost_Model_L_PSH_J16 = xl_mul((xl_eq(inputs.L_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_C45, "Horizontal") ? horizontal_Cost_Curves_L_PSH_P88 : (xl_eq(inputs.L_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_C45, "Vertical") ? vertical_Cost_Curves_L_PSH_P91 : missing)), 1.0e6) # Cost Model L-PSH J16
@assert xl_compare(s_Cost_Model_L_PSH_J16, 6.772919077968261e6) # "Cost Model L-PSH!J16"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[2, "N"])]
s_Cost_Model_L_PSH_J17 = calculate_s_Cost_Model_L_PSH_J17(inputs.Lower_Dam_Crest_Length_Cost_Model_L_PSH_C19, inputs.Avg_Lower_Dam_Height_Cost_Model_L_PSH_C18, tab_Cost_Curves_L_PSH_J95_J96)
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[1, "N"])]
s_Cost_Model_L_PSH_J28 = calculate_s_Cost_Model_L_PSH_J28(tab_Cost_Model_L_PSH_C97_C105, tab_Cost_Model_L_PSH_C80_C86, inputs.Ac_Ft_to_Cu_Ft_Cost_Model_L_PSH_C26, inputs.Generation_Time_Cost_Model_L_PSH_C21, inputs.Pump_Time__Cost_Model_L_PSH_C76)
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[2, "N"])]
s_Cost_Model_L_PSH_J29 = calculate_s_Cost_Model_L_PSH_J29(tab_Cost_Model_L_PSH_C97_C105, s_Cost_Curves_L_PSH_D374, inputs.s_Market_Adj_Factors_H51, inputs.s_Market_Adj_Factors_H58, tab_Market_Adj_Factors_C45_C48)
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[3, "N"])]
s_Cost_Model_L_PSH_J30 = calculate_s_Cost_Model_L_PSH_J30(tab_Cost_Model_L_PSH_C97_C105, inputs.Power_Station_Cost_Model_L_PSH_C47, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10, Mean_Gross_Head__Cost_Model_L_PSH_C89)
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[2, "N"])]
# =IF(C47="Surface",0,'Cost Curves L-PSH'!D242)
s_Cost_Model_L_PSH_J34 = (xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Surface") ? 0.0 : s_Cost_Curves_L_PSH_D242) # Cost Model L-PSH J34
@assert xl_compare(s_Cost_Model_L_PSH_J34, 2555.8188374587603) # "Cost Model L-PSH!J34"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[1, "N"])]
tab_Cost_Model_L_PSH_K13_K14[1, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K13 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K13_K14[1, "K"], 1.0) # "Cost Model L-PSH!K13"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[2, "N"])]
tab_Cost_Model_L_PSH_K13_K14[2, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K14 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K13_K14[2, "K"], 1.0) # "Cost Model L-PSH!K14"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[1, "N"])]
tab_Cost_Model_L_PSH_K16_N17[1, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K16 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[1, "K"], 1.0) # "Cost Model L-PSH!K16"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[2, "N"])]
tab_Cost_Model_L_PSH_K16_N17[2, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K17 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[2, "K"], 1.0) # "Cost Model L-PSH!K17"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[1, "N"])]
tab_Cost_Model_L_PSH_K28_K30[1, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K28 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K28_K30[1, "K"], 1.0) # "Cost Model L-PSH!K28"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[2, "N"])]
tab_Cost_Model_L_PSH_K28_K30[2, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K29 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K28_K30[2, "K"], 1.0) # "Cost Model L-PSH!K29"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[3, "N"])]
tab_Cost_Model_L_PSH_K28_K30[3, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K30 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_K28_K30[3, "K"], 1.0) # "Cost Model L-PSH!K30"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[2, "N"])]
tab_Cost_Model_L_PSH_K33_N34[2, "K"] = xl_div(xl_vlookup(inputs.Location_Cost_Model_L_PSH_C9, tab_Locational_Adj_Factors_A3_B55[!, Between("A", "B")], 2.0, false), 100.0) # Cost Model L-PSH K34 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[2, "K"], 1.0) # "Cost Model L-PSH!K34"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_N8)]
# =IF(C51="Yes",'Market Adj Factors'!C45,1)
s_Cost_Model_L_PSH_L8 = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L8
@assert xl_compare(s_Cost_Model_L_PSH_L8, 2.4063214260550976) # "Cost Model L-PSH!L8"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[1, "N"])]
tab_Cost_Model_L_PSH_L13_M13[1, "L"] = tab_Market_Adj_Factors_C45_C48[1, "C"] # Cost Model L-PSH L13 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_L13_M13[1, "L"], 2.4063214260550976) # "Cost Model L-PSH!L13"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[2, "N"])]
# =IF(C51="Yes",'Market Adj Factors'!C45,1)
s_Cost_Model_L_PSH_L14 = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L14
@assert xl_compare(s_Cost_Model_L_PSH_L14, 2.4063214260550976) # "Cost Model L-PSH!L14"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[1, "N"])]
tab_Cost_Model_L_PSH_K16_N17[1, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L16 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[1, "L"], 2.4063214260550976) # "Cost Model L-PSH!L16"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[2, "N"])]
tab_Cost_Model_L_PSH_K16_N17[2, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L17 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[2, "L"], 2.4063214260550976) # "Cost Model L-PSH!L17"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[3, "N"])]
# =IF($C$51="Yes",'Market Adj Factors'!$C$45,1)
s_Cost_Model_L_PSH_L30 = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L30
@assert xl_compare(s_Cost_Model_L_PSH_L30, 2.4063214260550976) # "Cost Model L-PSH!L30"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[2, "N"])]
tab_Cost_Model_L_PSH_K33_N34[2, "L"] = (xl_eq(inputs.Inflation_Factor_Cost_Model_L_PSH_C51, "Yes") ? tab_Market_Adj_Factors_C45_C48[1, "C"] : 1.0) # Cost Model L-PSH L34 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[2, "L"], 2.4063214260550976) # "Cost Model L-PSH!L34"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[1, "N"])]
tab_Cost_Model_L_PSH_L13_M13[1, "M"] = inputs.Dams_spillways_diversions_emb # Cost Model L-PSH M13 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_L13_M13[1, "M"], 1.4) # "Cost Model L-PSH!M13"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N13_N14[2, "N"])]
# ='Market Adj Factors'!C32
s_Cost_Model_L_PSH_M14 = inputs.Upper_intake_per_outlet_vertical # Cost Model L-PSH M14
@assert xl_compare(s_Cost_Model_L_PSH_M14, 2.0) # "Cost Model L-PSH!M14"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[1, "N"])]
tab_Cost_Model_L_PSH_K16_N17[1, "M"] = inputs.Lower_intake_per_outlet_horizontal # Cost Model L-PSH M16 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[1, "M"], 1.3) # "Cost Model L-PSH!M16"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K16_N17[2, "N"])]
tab_Cost_Model_L_PSH_K16_N17[2, "M"] = inputs.Dams_spillways_diversions_emb # Cost Model L-PSH M17 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[2, "M"], 1.4) # "Cost Model L-PSH!M17"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_N28_N30[3, "N"])]
# ='Market Adj Factors'!C41
s_Cost_Model_L_PSH_M30 = inputs.Electro_Mechanical_ # Cost Model L-PSH M30
@assert xl_compare(s_Cost_Model_L_PSH_M30, 1.7) # "Cost Model L-PSH!M30"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_K33_N34[2, "N"])]
tab_Cost_Model_L_PSH_K33_N34[2, "M"] = inputs.Access_and_voltage_tunnels # Cost Model L-PSH M34 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[2, "M"], 1.0) # "Cost Model L-PSH!M34"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[1, "Q"])]
tab_Cost_Model_L_PSH_K20_N25[1, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P20) ? inputs.s_Cost_Model_L_PSH_P20 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J20, tab_Cost_Model_L_PSH_K20_N25[1, "K"]), tab_Cost_Model_L_PSH_K20_N25[1, "L"]), tab_Cost_Model_L_PSH_K20_N25[1, "M"])) # Cost Model L-PSH N20 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[1, "N"], 14119.47764492283) # "Cost Model L-PSH!N20"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[2, "Q"])]
tab_Cost_Model_L_PSH_K20_N25[2, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P21) ? inputs.s_Cost_Model_L_PSH_P21 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J21, tab_Cost_Model_L_PSH_K20_N25[2, "K"]), tab_Cost_Model_L_PSH_K20_N25[2, "L"]), tab_Cost_Model_L_PSH_K20_N25[2, "M"])) # Cost Model L-PSH N21 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[2, "N"], 20624.348881754704) # "Cost Model L-PSH!N21"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[3, "Q"])]
tab_Cost_Model_L_PSH_K20_N25[3, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P22) ? inputs.s_Cost_Model_L_PSH_P22 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J22, tab_Cost_Model_L_PSH_K20_N25[3, "K"]), tab_Cost_Model_L_PSH_K20_N25[3, "L"]), tab_Cost_Model_L_PSH_K20_N25[3, "M"])) # Cost Model L-PSH N22 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[3, "N"], 22063.678916138877) # "Cost Model L-PSH!N22"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[4, "Q"])]
tab_Cost_Model_L_PSH_K20_N25[4, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P23) ? inputs.s_Cost_Model_L_PSH_P23 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J23, tab_Cost_Model_L_PSH_K20_N25[4, "K"]), tab_Cost_Model_L_PSH_K20_N25[4, "L"]), tab_Cost_Model_L_PSH_K20_N25[4, "M"])) # Cost Model L-PSH N23 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[4, "N"], 30175.534402399684) # "Cost Model L-PSH!N23"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[5, "Q"])]
tab_Cost_Model_L_PSH_K20_N25[5, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P24) ? inputs.s_Cost_Model_L_PSH_P24 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J24, tab_Cost_Model_L_PSH_K20_N25[5, "K"]), tab_Cost_Model_L_PSH_K20_N25[5, "L"]), tab_Cost_Model_L_PSH_K20_N25[5, "M"])) # Cost Model L-PSH N24 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[5, "N"], 14119.47764492283) # "Cost Model L-PSH!N24"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q20_R25[6, "Q"])]
tab_Cost_Model_L_PSH_K20_N25[6, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P25) ? inputs.s_Cost_Model_L_PSH_P25 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J25, tab_Cost_Model_L_PSH_K20_N25[6, "K"]), tab_Cost_Model_L_PSH_K20_N25[6, "L"]), tab_Cost_Model_L_PSH_K20_N25[6, "M"])) # Cost Model L-PSH N25 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_K20_N25[6, "N"], 0.0) # "Cost Model L-PSH!N25"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q33_R35[1, "Q"])]
tab_Cost_Model_L_PSH_K33_N34[1, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P33) ? inputs.s_Cost_Model_L_PSH_P33 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J33, tab_Cost_Model_L_PSH_K33_N34[1, "K"]), tab_Cost_Model_L_PSH_K33_N34[1, "L"]), tab_Cost_Model_L_PSH_K33_N34[1, "M"])) # Cost Model L-PSH N33 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[1, "N"], 454794.74952441343) # "Cost Model L-PSH!N33"


# Level 9
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_Q8)]
# =IF(G8="Yes",IF(O8,O8,C20),0)
Acres_Cost_Model_L_PSH_I8 = (xl_eq(inputs.Land_and_Land_Rights_Cost_Model_L_PSH_G8, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O8) ? inputs.s_Cost_Model_L_PSH_O8 : inputs.Acreage_to_be_acquired_Cost_Model_L_PSH_C20) : 0.0) # Cost Model L-PSH I8
@assert xl_compare(Acres_Cost_Model_L_PSH_I8, 2185) # "Cost Model L-PSH!I8"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q13_Q14[!, "Q"])]
tab_Cost_Model_L_PSH_I13_I14[1, "I"] = (xl_eq(inputs.Upper_Reservoir_Dam_and_Spillway_Cost_Model_L_PSH_G13, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O13) ? inputs.s_Cost_Model_L_PSH_O13 : tab_Cost_Model_L_PSH_C80_C86[4, "C"]) : 0.0) # Cost Model L-PSH I13 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_I13_I14[1, "I"], 2385110) # "Cost Model L-PSH!I13"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q13_Q14[!, "Q"])]
tab_Cost_Model_L_PSH_I13_I14[2, "I"] = (xl_eq(inputs.Upper_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_G14, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O14) ? inputs.s_Cost_Model_L_PSH_O14 : No_Tunnels__Cost_Model_L_PSH_C93) : 0.0) # Cost Model L-PSH I14 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_I13_I14[2, "I"], 1) # "Cost Model L-PSH!I14"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"])]
# =IF(G16="Yes",IF(C47="Underground",IF(O16,O16,C93),0))
num_Cost_Model_L_PSH_I16 = (xl_eq(inputs.Lower_Reservoir_Intake_per_Outlet_Cost_Model_L_PSH_G16, "Yes") ? (xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Underground") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O16) ? inputs.s_Cost_Model_L_PSH_O16 : No_Tunnels__Cost_Model_L_PSH_C93) : 0.0) : missing) # Cost Model L-PSH I16
@assert xl_compare(num_Cost_Model_L_PSH_I16, 1) # "Cost Model L-PSH!I16"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"])]
# =IF(G17="Yes",IF(O17,O17,C84),0)
CY_Cost_Model_L_PSH_I17 = (xl_eq(inputs.Lower_Reservoir_Dam_and_Spillway_Cost_Model_L_PSH_G17, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O17) ? inputs.s_Cost_Model_L_PSH_O17 : tab_Cost_Model_L_PSH_C80_C86[5, "C"]) : 0.0) # Cost Model L-PSH I17
@assert xl_compare(CY_Cost_Model_L_PSH_I17, 691570) # "Cost Model L-PSH!I17"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q28_R30[1, "Q"])]
# =IF(G28="Yes",1,0)
LS_Cost_Model_L_PSH_I28 = (xl_eq(inputs.Pump_per_Motors_Cost_Model_L_PSH_G28, "Yes") ? 1.0 : 0.0) # Cost Model L-PSH I28
@assert xl_compare(LS_Cost_Model_L_PSH_I28, 1) # "Cost Model L-PSH!I28"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q28_R30[2, "Q"])]
# =IF(G29="Yes",IF(C104<=100,IF(O29,O29,I10),0),0)
kW_Cost_Model_L_PSH_I29 = (xl_eq(inputs.Generator_per_Turbines_Cost_Model_L_PSH_G29, "Yes") ? (xl_leq(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 100.0) ? (xl_logical(inputs.s_Cost_Model_L_PSH_O29) ? inputs.s_Cost_Model_L_PSH_O29 : kW_Cost_Model_L_PSH_I10) : 0.0) : 0.0) # Cost Model L-PSH I29
@assert xl_compare(kW_Cost_Model_L_PSH_I29, 0) # "Cost Model L-PSH!I29"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q28_R30[3, "Q"])]
# =IF(G30="Yes",IF(C104>100,IF(O30,O30,C105*1000),0),0)
kW_Cost_Model_L_PSH_I30 = (xl_eq(inputs.Total_Powerstation_Cost_Model_L_PSH_G30, "Yes") ? (xl_gt(tab_Cost_Model_L_PSH_C97_C105[8, "C"], 100.0) ? (xl_logical(inputs.s_Cost_Model_L_PSH_O30) ? inputs.s_Cost_Model_L_PSH_O30 : tab_Cost_Model_L_PSH_C97_C105[9, "C"] * 1000.0) : 0.0) : 0.0) # Cost Model L-PSH I30
@assert xl_compare(kW_Cost_Model_L_PSH_I30, 1.2833248207034676e6) # "Cost Model L-PSH!I30"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q33_R35[2, "Q"])]
# =IF(G34="Yes",IF(C47="Surface",0,IF(O34,O34,C39*5280)),0)
Ft_Cost_Model_L_PSH_I34 = (xl_eq(inputs.Access_Tunnels_Cost_Model_L_PSH_G34, "Yes") ? (xl_eq(inputs.Power_Station_Cost_Model_L_PSH_C47, "Surface") ? 0.0 : (xl_logical(inputs.s_Cost_Model_L_PSH_O34) ? inputs.s_Cost_Model_L_PSH_O34 : inputs.Access_Tunnel_Length_Cost_Model_L_PSH_C39 * 5280.0)) : 0.0) # Cost Model L-PSH I34
@assert xl_compare(Ft_Cost_Model_L_PSH_I34, 6600) # "Cost Model L-PSH!I34"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q33_R35[3, "Q"])]
# =IF(G35="Yes",IF(C37="Yes",25%,0%),0)
pcnt_Cost_Model_L_PSH_I35 = (xl_eq(inputs.Highway_Realignment_Cost_Model_L_PSH_G35, "Yes") ? (xl_eq(inputs.Highway_Realignment_Cost_Model_L_PSH_C37, "Yes") ? ((25.0) / 100.0) : ((0.0) / 100.0)) : 0.0) # Cost Model L-PSH I35
@assert xl_compare(pcnt_Cost_Model_L_PSH_I35, 0.25) # "Cost Model L-PSH!I35"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_Q39)]
# =IF(G39="Yes",C61,0)
Miles_Cost_Model_L_PSH_I39 = (xl_eq(inputs.Transmission_Lines_Cost_Model_L_PSH_G39, "Yes") ? inputs.Transmission_Distance__Cost_Model_L_PSH_C61 : 0.0) # Cost Model L-PSH I39
@assert xl_compare(Miles_Cost_Model_L_PSH_I39, 13.5) # "Cost Model L-PSH!I39"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_Q8)]
# =IF(P8,P8,J8*K8*L8*M8)
s_Cost_Model_L_PSH_N8 = (xl_logical(inputs.s_Cost_Model_L_PSH_P8) ? inputs.s_Cost_Model_L_PSH_P8 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J8, inputs.s_Cost_Model_L_PSH_K8), s_Cost_Model_L_PSH_L8), inputs.s_Cost_Model_L_PSH_M8)) # Cost Model L-PSH N8
@assert xl_compare(s_Cost_Model_L_PSH_N8, 7443.554277930435) # "Cost Model L-PSH!N8"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q13_Q14[!, "Q"])]
tab_Cost_Model_L_PSH_N13_N14[1, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P13) ? inputs.s_Cost_Model_L_PSH_P13 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J13, tab_Cost_Model_L_PSH_K13_K14[1, "K"]), tab_Cost_Model_L_PSH_L13_M13[1, "L"]), tab_Cost_Model_L_PSH_L13_M13[1, "M"])) # Cost Model L-PSH N13 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_N13_N14[1, "N"], 26.669979831356148) # "Cost Model L-PSH!N13"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q13_Q14[!, "Q"])]
tab_Cost_Model_L_PSH_N13_N14[2, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P14) ? inputs.s_Cost_Model_L_PSH_P14 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J14, tab_Cost_Model_L_PSH_K13_K14[2, "K"]), s_Cost_Model_L_PSH_L14), s_Cost_Model_L_PSH_M14)) # Cost Model L-PSH N14 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_N13_N14[2, "N"], 7.138076300910789e6) # "Cost Model L-PSH!N14"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"])]
tab_Cost_Model_L_PSH_K16_N17[1, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P16) ? inputs.s_Cost_Model_L_PSH_P16 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J16, tab_Cost_Model_L_PSH_K16_N17[1, "K"]), tab_Cost_Model_L_PSH_K16_N17[1, "L"]), tab_Cost_Model_L_PSH_K16_N17[1, "M"])) # Cost Model L-PSH N16 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[1, "N"], 2.118716638252807e7) # "Cost Model L-PSH!N16"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"])]
tab_Cost_Model_L_PSH_K16_N17[2, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P17) ? inputs.s_Cost_Model_L_PSH_P17 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J17, tab_Cost_Model_L_PSH_K16_N17[2, "K"]), tab_Cost_Model_L_PSH_K16_N17[2, "L"]), tab_Cost_Model_L_PSH_K16_N17[2, "M"])) # Cost Model L-PSH N17 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K16_N17[2, "N"], 28.132802819876403) # "Cost Model L-PSH!N17"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q28_R30[1, "Q"])]
tab_Cost_Model_L_PSH_N28_N30[1, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P28) ? inputs.s_Cost_Model_L_PSH_P28 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J28, tab_Cost_Model_L_PSH_K28_K30[1, "K"]), inputs.s_Cost_Model_L_PSH_L28), inputs.s_Cost_Model_L_PSH_M28)) # Cost Model L-PSH N28 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_N28_N30[1, "N"], 0.0) # "Cost Model L-PSH!N28"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q28_R30[2, "Q"])]
tab_Cost_Model_L_PSH_N28_N30[2, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P29) ? inputs.s_Cost_Model_L_PSH_P29 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J29, tab_Cost_Model_L_PSH_K28_K30[2, "K"]), inputs.s_Cost_Model_L_PSH_L29), inputs.s_Cost_Model_L_PSH_M29)) # Cost Model L-PSH N29 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_N28_N30[2, "N"], 0.0) # "Cost Model L-PSH!N29"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q28_R30[3, "Q"])]
tab_Cost_Model_L_PSH_N28_N30[3, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P30) ? inputs.s_Cost_Model_L_PSH_P30 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J30, tab_Cost_Model_L_PSH_K28_K30[3, "K"]), s_Cost_Model_L_PSH_L30), s_Cost_Model_L_PSH_M30)) # Cost Model L-PSH N30 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_N28_N30[3, "N"], 422.2254104043288) # "Cost Model L-PSH!N30"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q33_R35[2, "Q"])]
tab_Cost_Model_L_PSH_K33_N34[2, "N"] = (xl_logical(inputs.s_Cost_Model_L_PSH_P34) ? inputs.s_Cost_Model_L_PSH_P34 : xl_mul(xl_mul(xl_mul(s_Cost_Model_L_PSH_J34, tab_Cost_Model_L_PSH_K33_N34[2, "K"]), tab_Cost_Model_L_PSH_K33_N34[2, "L"]), tab_Cost_Model_L_PSH_K33_N34[2, "M"])) # Cost Model L-PSH N34 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_K33_N34[2, "N"], 6150.121629692246) # "Cost Model L-PSH!N34"
# Used in 1 places: [StandardStatement(lhs = s_Cost_Model_L_PSH_Q39)]
s_Cost_Model_L_PSH_N39 = calculate_s_Cost_Model_L_PSH_N39(inputs.s_Cost_Model_L_PSH_P39, tab_Market_Adj_Factors_C45_C48, inputs.Inflation_Factor_Cost_Model_L_PSH_C51, inputs.Transmission_works, tab_Locational_Adj_Factors_A3_B55, inputs.Location_Cost_Model_L_PSH_C9, inputs.Transmission__Cost_Model_L_PSH_C62, inputs.Transmission_Type_num_circuits__Cost_Model_L_PSH_C65, tab_Cost_Curves_L_PSH_H276_I284, inputs.Transmission_Terrain___Cost_Model_L_PSH_C64, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, tab_Cost_Model_L_PSH_C97_C105, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10, inputs.s_Cost_Curves_L_PSH_J292, inputs.s_Cost_Curves_L_PSH_I292, inputs.s_Cost_Curves_L_PSH_I291, inputs.s_Cost_Curves_L_PSH_J291, inputs.s_Cost_Curves_L_PSH_I290, inputs.s_Cost_Curves_L_PSH_J290, tab_Cost_Curves_L_PSH_K290_K292)
# Used in 8 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., FunctionStatement(lhs = NA_Cost_Model_L_PSH_Q15)]
tab_Cost_Model_L_PSH_Q20_R25[1, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K20_N25[1, "N"], Ft_Cost_Model_L_PSH_I20) # Cost Model L-PSH Q20 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], 8.850441574778754e7) # "Cost Model L-PSH!Q20"
# Used in 8 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., FunctionStatement(lhs = NA_Cost_Model_L_PSH_Q15)]
tab_Cost_Model_L_PSH_Q20_R25[2, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K20_N25[2, "N"], Ft_Cost_Model_L_PSH_I21) # Cost Model L-PSH Q21 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], 2.7347886617206737e7) # "Cost Model L-PSH!Q21"
# Used in 8 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., FunctionStatement(lhs = NA_Cost_Model_L_PSH_Q15)]
tab_Cost_Model_L_PSH_Q20_R25[3, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K20_N25[3, "N"], Ft_Cost_Model_L_PSH_I22) # Cost Model L-PSH Q22 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], 2.925643824280015e7) # "Cost Model L-PSH!Q22"
# Used in 8 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., FunctionStatement(lhs = NA_Cost_Model_L_PSH_Q15)]
tab_Cost_Model_L_PSH_Q20_R25[4, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K20_N25[4, "N"], Ft_Cost_Model_L_PSH_I23) # Cost Model L-PSH Q23 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], 2.4140427521919746e7) # "Cost Model L-PSH!Q23"
# Used in 8 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., FunctionStatement(lhs = NA_Cost_Model_L_PSH_Q15)]
tab_Cost_Model_L_PSH_Q20_R25[5, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K20_N25[5, "N"], Ft_Cost_Model_L_PSH_I24) # Cost Model L-PSH Q24 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], 8.850441574778754e7) # "Cost Model L-PSH!Q24"
# Used in 8 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., FunctionStatement(lhs = NA_Cost_Model_L_PSH_Q15)]
tab_Cost_Model_L_PSH_Q20_R25[6, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K20_N25[6, "N"], Ft_Cost_Model_L_PSH_I25) # Cost Model L-PSH Q25 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], 0) # "Cost Model L-PSH!Q25"
# Used in 8 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., TableStatement(lhs = tab_Cost_Model_L_PSH_Q33_R35[3, "Q"])]
tab_Cost_Model_L_PSH_Q33_R35[1, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K33_N34[1, "N"], Miles_Cost_Model_L_PSH_I33) # Cost Model L-PSH Q33 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], 1.6418090457831325e6) # "Cost Model L-PSH!Q33"


# Level 10
# Used in 1 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45)]
tab_Cost_Model_L_PSH_I45_I50[1, "I"] = (xl_eq(inputs.Mobilization_per_Demobilization_Cost_Model_L_PSH_G45, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O45) ? inputs.s_Cost_Model_L_PSH_O45 : inputs.Mobilization_per_Demobilization__Cost_Model_L_PSH_C52) : 0.0) # Cost Model L-PSH I45 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_I45_I50[1, "I"], 0.05) # "Cost Model L-PSH!I45"
# Used in 1 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_I45_I50[2, "I"] = (xl_eq(inputs.Sales_Tax_Cost_Model_L_PSH_G46, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O46) ? inputs.s_Cost_Model_L_PSH_O46 : inputs.Sales_Tax__Cost_Model_L_PSH_C54) : 0.0) # Cost Model L-PSH I46 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_I45_I50[2, "I"], 0.06) # "Cost Model L-PSH!I46"
# Used in 1 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47)]
tab_Cost_Model_L_PSH_I45_I50[3, "I"] = (xl_eq(inputs.Contingency_Cost_Model_L_PSH_G47, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O47) ? inputs.s_Cost_Model_L_PSH_O47 : inputs.Contingency__Cost_Model_L_PSH_C55) : 0.0) # Cost Model L-PSH I47 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_I45_I50[3, "I"], 0.33) # "Cost Model L-PSH!I47"
# Used in 1 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48)]
tab_Cost_Model_L_PSH_I45_I50[4, "I"] = (xl_eq(inputs.EPC_Cost_Cost_Model_L_PSH_G48, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O48) ? inputs.s_Cost_Model_L_PSH_O48 : inputs.EPC_Cost__Cost_Model_L_PSH_C56) : 0.0) # Cost Model L-PSH I48 Row: 4
@assert xl_compare(tab_Cost_Model_L_PSH_I45_I50[4, "I"], 0.25) # "Cost Model L-PSH!I48"
# Used in 1 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49)]
tab_Cost_Model_L_PSH_I45_I50[5, "I"] = (xl_eq(inputs.Developer_Cost_Cost_Model_L_PSH_G49, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O49) ? inputs.s_Cost_Model_L_PSH_O49 : inputs.Developer_Cost__Cost_Model_L_PSH_C57) : 0.0) # Cost Model L-PSH I49 Row: 5
@assert xl_compare(tab_Cost_Model_L_PSH_I45_I50[5, "I"], 0.03) # "Cost Model L-PSH!I49"
# Used in 1 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q50)]
tab_Cost_Model_L_PSH_I45_I50[6, "I"] = (xl_eq(inputs.Overhead__and__Profit_Cost_Model_L_PSH_G50, "Yes") ? (xl_logical(inputs.s_Cost_Model_L_PSH_O50) ? inputs.s_Cost_Model_L_PSH_O50 : inputs.Overhead__and__Profit__Cost_Model_L_PSH_C58) : 0.0) # Cost Model L-PSH I50 Row: 6
@assert xl_compare(tab_Cost_Model_L_PSH_I45_I50[6, "I"], 0.07) # "Cost Model L-PSH!I50"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
# =I8*N8
s_Cost_Model_L_PSH_Q8 = xl_mul(Acres_Cost_Model_L_PSH_I8, s_Cost_Model_L_PSH_N8) # Cost Model L-PSH Q8
@assert xl_compare(s_Cost_Model_L_PSH_Q8, 1.6264166097278e7) # "Cost Model L-PSH!Q8"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
# =N10*I10
s_Cost_Model_L_PSH_Q10 = xl_mul(s_Cost_Model_L_PSH_N10, kW_Cost_Model_L_PSH_I10) # Cost Model L-PSH Q10
@assert xl_compare(s_Cost_Model_L_PSH_Q10, 1.5680624794734025e8) # "Cost Model L-PSH!Q10"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
# "Cost Model L-PSH!Q13":"Cost Model L-PSH!Q14"
@. tab_Cost_Model_L_PSH_Q13_Q14[!, "Q"] = xl_mul(tab_Cost_Model_L_PSH_N13_N14[!, "N"], tab_Cost_Model_L_PSH_I13_I14[!, "I"])
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
NA_Cost_Model_L_PSH_Q15 = calculate_NA_Cost_Model_L_PSH_Q15(tab_Cost_Model_L_PSH_Q20_R25, inputs.NA_Cost_Model_L_PSH_O15, inputs.Surge_Facilities_Cost_Model_L_PSH_G15, Mean_Gross_Head__Cost_Model_L_PSH_C89, inputs.Total_Conveyance_Length_vert_plus_horiz_Cost_Model_L_PSH_C15)
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K16_N17[1, "N"], num_Cost_Model_L_PSH_I16) # Cost Model L-PSH Q16 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], 2.118716638252807e7) # "Cost Model L-PSH!Q16"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K16_N17[2, "N"], CY_Cost_Model_L_PSH_I17) # Cost Model L-PSH Q17 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], 1.9455802446141925e7) # "Cost Model L-PSH!Q17"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_Q28_R30[1, "Q"] = xl_mul(tab_Cost_Model_L_PSH_N28_N30[1, "N"], LS_Cost_Model_L_PSH_I28) # Cost Model L-PSH Q28 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], 0) # "Cost Model L-PSH!Q28"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_Q28_R30[2, "Q"] = xl_mul(tab_Cost_Model_L_PSH_N28_N30[2, "N"], kW_Cost_Model_L_PSH_I29) # Cost Model L-PSH Q29 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], 0) # "Cost Model L-PSH!Q29"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_Q28_R30[3, "Q"] = xl_mul(tab_Cost_Model_L_PSH_N28_N30[3, "N"], kW_Cost_Model_L_PSH_I30) # Cost Model L-PSH Q30 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], 5.418523491035832e8) # "Cost Model L-PSH!Q30"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_Q33_R35[2, "Q"] = xl_mul(tab_Cost_Model_L_PSH_K33_N34[2, "N"], Ft_Cost_Model_L_PSH_I34) # Cost Model L-PSH Q34 Row: 2
@assert xl_compare(tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], 4.0590802755968824e7) # "Cost Model L-PSH!Q34"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
tab_Cost_Model_L_PSH_Q33_R35[3, "Q"] = xl_mul(tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], pcnt_Cost_Model_L_PSH_I35) # Cost Model L-PSH Q35 Row: 3
@assert xl_compare(tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], 410452.26144578314) # "Cost Model L-PSH!Q35"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
s_Cost_Model_L_PSH_Q37 = calculate_s_Cost_Model_L_PSH_Q37(inputs.Switchyard_Cost_Model_L_PSH_G37, inputs.s_Cost_Model_L_PSH_P37, tab_Locational_Adj_Factors_A3_B55, inputs.Location_Cost_Model_L_PSH_C9, tab_Market_Adj_Factors_C45_C48, inputs.Inflation_Factor_Cost_Model_L_PSH_C51, inputs.Switchyard_Market_Adj_Factors_C43, inputs.Substation__Cost_Model_L_PSH_C63, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, tab_Cost_Model_L_PSH_C97_C105, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10)
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
# =N39*I39
s_Cost_Model_L_PSH_Q39 = xl_mul(s_Cost_Model_L_PSH_N39, Miles_Cost_Model_L_PSH_I39) # Cost Model L-PSH Q39
@assert xl_compare(s_Cost_Model_L_PSH_Q39, 4.8578635297828086e7) # "Cost Model L-PSH!Q39"
# Used in 7 places: [StandardStatement(lhs = NA_Cost_Model_L_PSH_Q45), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q49), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q48), StandardStatement(lhs = NA_Cost_Model_L_PSH_Q47), ..., StandardStatement(lhs = NA_Cost_Model_L_PSH_Q46)]
s_Cost_Model_L_PSH_Q42 = calculate_s_Cost_Model_L_PSH_Q42(inputs.s_Cost_Model_L_PSH_L42, inputs.s_Cost_Model_L_PSH_M42, inputs.s_Cost_Model_L_PSH_P42, inputs.Water_Supply__Cost_Model_L_PSH_C40, inputs.Water_Supply_Cost_Model_L_PSH_G42, Mean_Gen_Discharge__Cost_Model_L_PSH_C87, tab_Cost_Model_L_PSH_C90_C91, Nominal_Tunnel_Dia__Cost_Model_L_PSH_C92, No_Tunnels__Cost_Model_L_PSH_C93, Adjusted_Tunnel_Dia__Cost_Model_L_PSH_C94, tab_Cost_Model_L_PSH_C97_C105, No_Units__Cost_Model_L_PSH_C106, Unit_Rating__Cost_Model_L_PSH_C107, s_Cost_Model_L_PSH_J10, s_Cost_Model_L_PSH_N10, inputs.Water_Supply_Cost___Cost_Model_L_PSH_C41, tab_Locational_Adj_Factors_A3_B55, inputs.Location_Cost_Model_L_PSH_C9)


# Level 11
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q52_S52[1, "Q"])]
# =I45*SUM(Q8:Q42)
NA_Cost_Model_L_PSH_Q45 = tab_Cost_Model_L_PSH_I45_I50[1, "I"] * xl_sum([s_Cost_Model_L_PSH_Q8, inputs.s_Cost_Model_L_PSH_Q9, s_Cost_Model_L_PSH_Q10, inputs.s_Cost_Model_L_PSH_Q11, inputs.s_Cost_Model_L_PSH_Q12, tab_Cost_Model_L_PSH_Q13_Q14[1, "Q"], tab_Cost_Model_L_PSH_Q13_Q14[2, "Q"], NA_Cost_Model_L_PSH_Q15, tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], inputs.s_Cost_Model_L_PSH_Q18, inputs.s_Cost_Model_L_PSH_Q19, tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], inputs.s_Cost_Model_L_PSH_Q26, inputs.s_Cost_Model_L_PSH_Q27, tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], inputs.s_Cost_Model_L_PSH_Q31, inputs.s_Cost_Model_L_PSH_Q32, tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], inputs.s_Cost_Model_L_PSH_Q36, s_Cost_Model_L_PSH_Q37, inputs.s_Cost_Model_L_PSH_Q38, s_Cost_Model_L_PSH_Q39, inputs.s_Cost_Model_L_PSH_Q40, inputs.s_Cost_Model_L_PSH_Q41, s_Cost_Model_L_PSH_Q42]) # Cost Model L-PSH Q45
@assert xl_compare(NA_Cost_Model_L_PSH_Q45, 6.593207089461073e7) # "Cost Model L-PSH!Q45"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q52_S52[1, "Q"])]
# =I46*C53*SUM(Q8:Q42)
NA_Cost_Model_L_PSH_Q46 = tab_Cost_Model_L_PSH_I45_I50[2, "I"] * inputs.Material_per_Equipment_pcnt__Cost_Model_L_PSH_C53 * xl_sum([s_Cost_Model_L_PSH_Q8, inputs.s_Cost_Model_L_PSH_Q9, s_Cost_Model_L_PSH_Q10, inputs.s_Cost_Model_L_PSH_Q11, inputs.s_Cost_Model_L_PSH_Q12, tab_Cost_Model_L_PSH_Q13_Q14[1, "Q"], tab_Cost_Model_L_PSH_Q13_Q14[2, "Q"], NA_Cost_Model_L_PSH_Q15, tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], inputs.s_Cost_Model_L_PSH_Q18, inputs.s_Cost_Model_L_PSH_Q19, tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], inputs.s_Cost_Model_L_PSH_Q26, inputs.s_Cost_Model_L_PSH_Q27, tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], inputs.s_Cost_Model_L_PSH_Q31, inputs.s_Cost_Model_L_PSH_Q32, tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], inputs.s_Cost_Model_L_PSH_Q36, s_Cost_Model_L_PSH_Q37, inputs.s_Cost_Model_L_PSH_Q38, s_Cost_Model_L_PSH_Q39, inputs.s_Cost_Model_L_PSH_Q40, inputs.s_Cost_Model_L_PSH_Q41, s_Cost_Model_L_PSH_Q42]) # Cost Model L-PSH Q46
@assert xl_compare(NA_Cost_Model_L_PSH_Q46, 7.911848507353286e7) # "Cost Model L-PSH!Q46"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q52_S52[1, "Q"])]
# =SUM(Q8:Q42)*I47
NA_Cost_Model_L_PSH_Q47 = xl_sum([s_Cost_Model_L_PSH_Q8, inputs.s_Cost_Model_L_PSH_Q9, s_Cost_Model_L_PSH_Q10, inputs.s_Cost_Model_L_PSH_Q11, inputs.s_Cost_Model_L_PSH_Q12, tab_Cost_Model_L_PSH_Q13_Q14[1, "Q"], tab_Cost_Model_L_PSH_Q13_Q14[2, "Q"], NA_Cost_Model_L_PSH_Q15, tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], inputs.s_Cost_Model_L_PSH_Q18, inputs.s_Cost_Model_L_PSH_Q19, tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], inputs.s_Cost_Model_L_PSH_Q26, inputs.s_Cost_Model_L_PSH_Q27, tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], inputs.s_Cost_Model_L_PSH_Q31, inputs.s_Cost_Model_L_PSH_Q32, tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], inputs.s_Cost_Model_L_PSH_Q36, s_Cost_Model_L_PSH_Q37, inputs.s_Cost_Model_L_PSH_Q38, s_Cost_Model_L_PSH_Q39, inputs.s_Cost_Model_L_PSH_Q40, inputs.s_Cost_Model_L_PSH_Q41, s_Cost_Model_L_PSH_Q42]) * tab_Cost_Model_L_PSH_I45_I50[3, "I"] # Cost Model L-PSH Q47
@assert xl_compare(NA_Cost_Model_L_PSH_Q47, 4.351516679044308e8) # "Cost Model L-PSH!Q47"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q52_S52[1, "Q"])]
# =I48*C53*SUM(Q8:Q42)
NA_Cost_Model_L_PSH_Q48 = tab_Cost_Model_L_PSH_I45_I50[4, "I"] * inputs.Material_per_Equipment_pcnt__Cost_Model_L_PSH_C53 * xl_sum([s_Cost_Model_L_PSH_Q8, inputs.s_Cost_Model_L_PSH_Q9, s_Cost_Model_L_PSH_Q10, inputs.s_Cost_Model_L_PSH_Q11, inputs.s_Cost_Model_L_PSH_Q12, tab_Cost_Model_L_PSH_Q13_Q14[1, "Q"], tab_Cost_Model_L_PSH_Q13_Q14[2, "Q"], NA_Cost_Model_L_PSH_Q15, tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], inputs.s_Cost_Model_L_PSH_Q18, inputs.s_Cost_Model_L_PSH_Q19, tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], inputs.s_Cost_Model_L_PSH_Q26, inputs.s_Cost_Model_L_PSH_Q27, tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], inputs.s_Cost_Model_L_PSH_Q31, inputs.s_Cost_Model_L_PSH_Q32, tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], inputs.s_Cost_Model_L_PSH_Q36, s_Cost_Model_L_PSH_Q37, inputs.s_Cost_Model_L_PSH_Q38, s_Cost_Model_L_PSH_Q39, inputs.s_Cost_Model_L_PSH_Q40, inputs.s_Cost_Model_L_PSH_Q41, s_Cost_Model_L_PSH_Q42]) # Cost Model L-PSH Q48
@assert xl_compare(NA_Cost_Model_L_PSH_Q48, 3.2966035447305363e8) # "Cost Model L-PSH!Q48"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q52_S52[1, "Q"])]
# =SUM(Q8:Q42)*I49
NA_Cost_Model_L_PSH_Q49 = xl_sum([s_Cost_Model_L_PSH_Q8, inputs.s_Cost_Model_L_PSH_Q9, s_Cost_Model_L_PSH_Q10, inputs.s_Cost_Model_L_PSH_Q11, inputs.s_Cost_Model_L_PSH_Q12, tab_Cost_Model_L_PSH_Q13_Q14[1, "Q"], tab_Cost_Model_L_PSH_Q13_Q14[2, "Q"], NA_Cost_Model_L_PSH_Q15, tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], inputs.s_Cost_Model_L_PSH_Q18, inputs.s_Cost_Model_L_PSH_Q19, tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], inputs.s_Cost_Model_L_PSH_Q26, inputs.s_Cost_Model_L_PSH_Q27, tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], inputs.s_Cost_Model_L_PSH_Q31, inputs.s_Cost_Model_L_PSH_Q32, tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], inputs.s_Cost_Model_L_PSH_Q36, s_Cost_Model_L_PSH_Q37, inputs.s_Cost_Model_L_PSH_Q38, s_Cost_Model_L_PSH_Q39, inputs.s_Cost_Model_L_PSH_Q40, inputs.s_Cost_Model_L_PSH_Q41, s_Cost_Model_L_PSH_Q42]) * tab_Cost_Model_L_PSH_I45_I50[5, "I"] # Cost Model L-PSH Q49
@assert xl_compare(NA_Cost_Model_L_PSH_Q49, 3.955924253676643e7) # "Cost Model L-PSH!Q49"
# Used in 1 places: [TableStatement(lhs = tab_Cost_Model_L_PSH_Q52_S52[1, "Q"])]
# =SUM(Q8:Q42)*I50*C53
NA_Cost_Model_L_PSH_Q50 = xl_sum([s_Cost_Model_L_PSH_Q8, inputs.s_Cost_Model_L_PSH_Q9, s_Cost_Model_L_PSH_Q10, inputs.s_Cost_Model_L_PSH_Q11, inputs.s_Cost_Model_L_PSH_Q12, tab_Cost_Model_L_PSH_Q13_Q14[1, "Q"], tab_Cost_Model_L_PSH_Q13_Q14[2, "Q"], NA_Cost_Model_L_PSH_Q15, tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], inputs.s_Cost_Model_L_PSH_Q18, inputs.s_Cost_Model_L_PSH_Q19, tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], inputs.s_Cost_Model_L_PSH_Q26, inputs.s_Cost_Model_L_PSH_Q27, tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], inputs.s_Cost_Model_L_PSH_Q31, inputs.s_Cost_Model_L_PSH_Q32, tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], inputs.s_Cost_Model_L_PSH_Q36, s_Cost_Model_L_PSH_Q37, inputs.s_Cost_Model_L_PSH_Q38, s_Cost_Model_L_PSH_Q39, inputs.s_Cost_Model_L_PSH_Q40, inputs.s_Cost_Model_L_PSH_Q41, s_Cost_Model_L_PSH_Q42]) * tab_Cost_Model_L_PSH_I45_I50[6, "I"] * inputs.Material_per_Equipment_pcnt__Cost_Model_L_PSH_C53 # Cost Model L-PSH Q50
@assert xl_compare(NA_Cost_Model_L_PSH_Q50, 9.230489925245503e7) # "Cost Model L-PSH!Q50"


# Level 12
# Used in 1 places: [StandardStatement(lhs = Total_Direct_and_Indirect_Cost__dollar_Cost_Model_L_PSH_G55)]
tab_Cost_Model_L_PSH_Q52_S52[1, "Q"] = xl_sum([s_Cost_Model_L_PSH_Q8, inputs.s_Cost_Model_L_PSH_Q9, s_Cost_Model_L_PSH_Q10, inputs.s_Cost_Model_L_PSH_Q11, inputs.s_Cost_Model_L_PSH_Q12, tab_Cost_Model_L_PSH_Q13_Q14[1, "Q"], tab_Cost_Model_L_PSH_Q13_Q14[2, "Q"], NA_Cost_Model_L_PSH_Q15, tab_Cost_Model_L_PSH_Q16_Q17[1, "Q"], tab_Cost_Model_L_PSH_Q16_Q17[2, "Q"], inputs.s_Cost_Model_L_PSH_Q18, inputs.s_Cost_Model_L_PSH_Q19, tab_Cost_Model_L_PSH_Q20_R25[1, "Q"], tab_Cost_Model_L_PSH_Q20_R25[2, "Q"], tab_Cost_Model_L_PSH_Q20_R25[3, "Q"], tab_Cost_Model_L_PSH_Q20_R25[4, "Q"], tab_Cost_Model_L_PSH_Q20_R25[5, "Q"], tab_Cost_Model_L_PSH_Q20_R25[6, "Q"], inputs.s_Cost_Model_L_PSH_Q26, inputs.s_Cost_Model_L_PSH_Q27, tab_Cost_Model_L_PSH_Q28_R30[1, "Q"], tab_Cost_Model_L_PSH_Q28_R30[2, "Q"], tab_Cost_Model_L_PSH_Q28_R30[3, "Q"], inputs.s_Cost_Model_L_PSH_Q31, inputs.s_Cost_Model_L_PSH_Q32, tab_Cost_Model_L_PSH_Q33_R35[1, "Q"], tab_Cost_Model_L_PSH_Q33_R35[2, "Q"], tab_Cost_Model_L_PSH_Q33_R35[3, "Q"], inputs.s_Cost_Model_L_PSH_Q36, s_Cost_Model_L_PSH_Q37, inputs.s_Cost_Model_L_PSH_Q38, s_Cost_Model_L_PSH_Q39, inputs.s_Cost_Model_L_PSH_Q40, inputs.s_Cost_Model_L_PSH_Q41, s_Cost_Model_L_PSH_Q42, inputs.s_Cost_Model_L_PSH_Q43, inputs.s_Cost_Model_L_PSH_Q44, NA_Cost_Model_L_PSH_Q45, NA_Cost_Model_L_PSH_Q46, NA_Cost_Model_L_PSH_Q47, NA_Cost_Model_L_PSH_Q48, NA_Cost_Model_L_PSH_Q49, NA_Cost_Model_L_PSH_Q50]) # Cost Model L-PSH Q52 Row: 1
@assert xl_compare(tab_Cost_Model_L_PSH_Q52_S52[1, "Q"], 2.3603681380270643e9) # "Cost Model L-PSH!Q52"


# Level 13
# Used in 1 places: [StandardStatement(lhs = _dollar_per_kWh_Max_Energy_Capacity_Cost_Model_L_PSH_G57)]
# =Q52
Total_Direct_and_Indirect_Cost__dollar_Cost_Model_L_PSH_G55 = tab_Cost_Model_L_PSH_Q52_S52[1, "Q"] # Cost Model L-PSH G55
@assert xl_compare(Total_Direct_and_Indirect_Cost__dollar_Cost_Model_L_PSH_G55, 2.3603681380270643e9) # "Cost Model L-PSH!G55"


# Level 14
# Used in 1 places: [OutputStatement]
# =G55/(C105*C21*1000)
_dollar_per_kWh_Max_Energy_Capacity_Cost_Model_L_PSH_G57 = Total_Direct_and_Indirect_Cost__dollar_Cost_Model_L_PSH_G55 / (tab_Cost_Model_L_PSH_C97_C105[9, "C"] * inputs.Generation_Time_Cost_Model_L_PSH_C21 * 1000.0) # Cost Model L-PSH G57
@assert xl_compare(_dollar_per_kWh_Max_Energy_Capacity_Cost_Model_L_PSH_G57, 99.4194648634969) # "Cost Model L-PSH!G57"


# Level 15
Outputs(
    _dollar_per_kWh_Max_Energy_Capacity_Cost_Model_L_PSH_G57    
)


end

function run_crest_solar()
    inputs = Inputs()
    tables = make_input_tables()
    calculate(inputs, tables)
end
