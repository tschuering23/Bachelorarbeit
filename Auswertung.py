import pandas as pd
import numpy as np

strombezugspreis = 0.2673
dieselbezugspreis = 0.1491
einspeiseverguetung_pv = 0.0958
einspeiseverguetung_wind = 0.0736
energycosts_statusquo = 2226960.1 #Laufende Energiekosten ohne Investition
min_invest = 312000 * 20

results = pd.read_csv("results.csv")
components = pd.read_csv("components.csv")
pv = components[components['ID'].str.contains("PV")]
lkw = components[components['ID'].str.contains("lkw")]
e_lkw = components[components['ID'].str.contains("e-lkw")]
wasserstoff_lkw = components[components['ID'].str.contains("wasserstoff-lkw")]

diesel_lkw = components[components['ID'].str.contains("diesel-lkw")]

wind_bus = pd.read_excel("results_ID_wind_bus.xlsx")
niederspannung_bus = pd.read_excel("results_ID_niederspannung_bus.xlsx")
netzanschluss_bus = pd.read_excel("results_ID_netzanschluss_bus.xlsx")
wasserstoff_bus = pd.read_excel("results_ID_wasserstoff_bus.xlsx")
pv_bus  = pd.read_excel("results_ID_pv_bus.xlsx")

ergebnisstabelle = pd.read_excel("Ergebnisse.xlsx")
components.set_index("ID", inplace = True)

#Gesamt 
invest_mehr = round(components["periodical costs/CU"].sum() * 20 + components["variable costs/CU"].sum() * 20 - min_invest,2)#
energycosts_gesamt = round(components["constraints/CU"].sum(),2)#
energycosts_reduction = energycosts_statusquo - energycosts_gesamt#
gewinn_nach_20_jahren = energycosts_reduction * 20 - invest_mehr
if invest_mehr == 0:
    amortisation = 0
    rentabilität = 0
else:
    amortisation = (invest_mehr) / energycosts_reduction#
    rentabilität = (gewinn_nach_20_jahren / invest_mehr)/20 * 100#

#Netzentgelt
leistung = components.loc["ID_Netzbezug_link"]["investment/kW"]
netzentgelt = components.loc["ID_Netzbezug_link"]["constraints/CU"]

#Anteil Erneuerbare an Wasserstoffproduktion, Batterie und E-Mobilität
'''
Zur internen Berechnung der Gestehungskosten, 
da Wasserstoffproduktion und E-Mobilität gleichzeitig verzicht auf Einspeisung bedeutet
Rückverstromung und Ausspeicherung aus Batteriespeicher vernachlässigt
'''
a = niederspannung_bus
a = a.assign(wind_kWh = wind_bus["(('ID_wind_bus', 'ID_wind__link'), 'flow')"])
a = a.assign(netz_kWh = netzanschluss_bus["(('ID_netzanschluss_bus_shortage', 'ID_netzanschluss_bus'), 'flow')"])
a = a.assign(pv_kWh = niederspannung_bus["(('ID_pv_link', 'ID_niederspannung_bus'), 'flow')"])
a = a.assign(gesamtenergie_kWh = a["pv_kWh"] + a["wind_kWh"] +a["netz_kWh"] )
a = a.assign(anteil_wind = a["wind_kWh"] / a["gesamtenergie_kWh"])
a = a.assign(anteil_pv = a["pv_kWh"] / a["gesamtenergie_kWh"])
a = a.assign(anteil_netz = (a["netz_kWh"]) / a["gesamtenergie_kWh"])

a = a.assign(wind_in_battery = a["anteil_wind"] * a["(('ID_niederspannung_bus', 'ID_battery_storage'), 'flow')"])
a = a.assign(pv_in_battery = a["anteil_pv"] * a["(('ID_niederspannung_bus', 'ID_battery_storage'), 'flow')"])
a = a.assign(netz_in_battery = a["anteil_netz"] * a["(('ID_niederspannung_bus', 'ID_battery_storage'), 'flow')"])


a = a.assign(wind_in_elektromobilität = a["anteil_wind"] * a["(('ID_niederspannung_bus', 'ID_e-ladesaeule_link'), 'flow')"])
a = a.assign(pv_in_elektromobilität = a["anteil_pv"] * a["(('ID_niederspannung_bus', 'ID_e-ladesaeule_link'), 'flow')"])
a = a.assign(netz_in_elektromobilität = a["anteil_netz"] * a["(('ID_niederspannung_bus', 'ID_e-ladesaeule_link'), 'flow')"])

a = a.assign(wind_in_elektrolyse = a["anteil_wind"] * a["(('ID_niederspannung_bus', 'ID_elektrolyse_transformer'), 'flow')"])
a = a.assign(pv_in_elektrolyse = a["anteil_pv"] * a["(('ID_niederspannung_bus', 'ID_elektrolyse_transformer'), 'flow')"])
a = a.assign(netz_in_elektrolyse = a["anteil_netz"] * a["(('ID_niederspannung_bus', 'ID_elektrolyse_transformer'), 'flow')"])


#PV
pv_kw = pv["investment/kW"].sum()#
pv_invest = pv["periodical costs/CU"].sum() * 20 + pv["variable costs/CU"].sum() * 20#
pv_energycosts = pv["constraints/CU"].sum()#
pv_energy_output = pv["output 1/kWh"].sum()#
pv_eigenverbrauch = pv_bus["(('ID_pv_bus', 'ID_pv_link'), 'flow')"].sum()
pv_einspeisung = pv_bus["(('ID_pv_bus', 'ID_pv_transformator_link'), 'flow')"].sum()
if pv_kw  == 0:
    pv_stromgehstehungskosten = 0.0
    pv_eigenverbrauch_prozentual = 0.0
else:
    pv_stromgehstehungskosten = (pv_invest/20 + pv_energycosts) / pv_energy_output#
    pv_eigenverbrauch_prozentual = pv_eigenverbrauch * 100 /pv_energy_output#



#Wind
wind_kw = components.loc["Windkraftanlage"]["investment/kW"]#
wind_invest = components.loc["Windkraftanlage"]["periodical costs/CU"] * 20 + components.loc["Windkraftanlage"]["variable costs/CU"] * 20#
wind_energycosts = components.loc["Windkraftanlage"]["constraints/CU"]#
    
wind_energy_output = wind_bus["(('Windkraftanlage', 'ID_wind_bus'), 'flow')"].sum()#
wind_eigenverbrauch = wind_bus["(('ID_wind_bus', 'ID_wind__link'), 'flow')"].sum()
wind_einspeisung = wind_bus["(('ID_wind_bus', 'ID_wind_bus_excess'), 'flow')"].sum()
if wind_kw == 0:
    wind_stromgehstehungskosten = 0.0
    wind_eigenverbrauch_prozentual = 0.0
else:
    wind_stromgehstehungskosten = (wind_invest/20 + wind_energycosts)/ wind_energy_output#
    wind_eigenverbrauch_prozentual = wind_eigenverbrauch * 100 /wind_energy_output#


#Erneuerbare
erneuerbar_invest = wind_invest + pv_invest#
erneuerbar_energycosts = wind_energycosts + pv_energycosts#    
erneuerbar_energy_output = pv_energy_output + wind_energy_output#
if erneuerbar_energy_output == 0:
    erneuerbar_stromgehstehungskosten = 0.0
    erneuerbar_eigenverbrauch_prozentual = 0.0
else:
    erneuerbar_stromgehstehungskosten = (erneuerbar_invest/20 + erneuerbar_energycosts)/ erneuerbar_energy_output#
    erneuerbar_eigenverbrauch_prozentual = (wind_eigenverbrauch + pv_eigenverbrauch) *100 / erneuerbar_energy_output#

#Batteriespeicher
battery_kwh = components.loc["ID_battery_storage"]["investment/kW"]#
battery_invest = components.loc["ID_battery_storage"]["periodical costs/CU"] * 20 + components.loc["ID_battery_storage"]["variable costs/CU"] * 20#
battery_energycosts = components.loc["ID_battery_storage"]["constraints/CU"]#
battery_energy_output = components.loc["ID_battery_storage"]["output 1/kWh"]#
battery_energy_input = components.loc["ID_battery_storage"]["input 1/kWh"]

wind_in_battery_summe = a["wind_in_battery"].sum()
pv_in_battery_summe = a["pv_in_battery"].sum()
netz_in_battery_summe = a["netz_in_battery"].sum()
battery_energycosts_strombezug = pv_in_battery_summe * einspeiseverguetung_pv + wind_in_battery_summe * einspeiseverguetung_wind + netz_in_battery_summe * strombezugspreis

if battery_kwh == 0:
    battery_stromgehstehungskosten = 0.0
else:
    battery_stromgehstehungskosten = (battery_invest/20 + battery_energycosts + battery_energycosts_strombezug)/ battery_energy_output#
if battery_kwh == 0:
    umschlagszahl = 0.0
else:
    umschlagszahl = battery_energy_output / battery_kwh#


#Wasserstoffproduktion
hydrogen = components.loc[["ID_wasserstoff_storage","ID_elektrolyse_transformer"]]  
elektrolyse_kw = hydrogen.loc["ID_elektrolyse_transformer"]["investment/kW"]
hydrogen_storage_kWh = hydrogen.loc["ID_wasserstoff_storage"]["investment/kW"]

hydrogen_invest = hydrogen["periodical costs/CU"].sum() * 20 + hydrogen["variable costs/CU"].sum() * 20
hydrogen_energycosts_anlage = hydrogen["constraints/CU"].sum()
hydrogen_output_kWh = components.loc["ID_elektrolyse_transformer"]["output 1/kWh"]
hydrogen_output_kg = hydrogen_output_kWh / 33.333

             
wind_in_H2_summe = a["wind_in_elektrolyse"].sum()
pv_in_H2_summe = a["pv_in_elektrolyse"].sum()
netz_in_H2_summe = a["netz_in_elektrolyse"].sum()


hydrogen_energycosts_strombezug = pv_in_H2_summe * einspeiseverguetung_pv + wind_in_H2_summe * einspeiseverguetung_wind + netz_in_H2_summe * strombezugspreis
hydrogen_energycosts_gesamt = hydrogen_energycosts_anlage + hydrogen_energycosts_strombezug
hydrogen_energycosts_gesamt = round(hydrogen_energycosts_gesamt,3)

if elektrolyse_kw == 0:
    hydrogengehstehungskosten_kWh = 0.0
    hydrogengehstehungskosten_kg = 0.0
    Erneuerbar_in_H2_anteil = 0.0
else:
    Erneuerbar_in_H2_anteil = (wind_in_H2_summe + pv_in_H2_summe) / a["(('ID_niederspannung_bus', 'ID_elektrolyse_transformer'), 'flow')"].sum()
    hydrogengehstehungskosten_kWh = ((hydrogen.loc["ID_elektrolyse_transformer"]["periodical costs/CU"] + hydrogen.loc["ID_elektrolyse_transformer"]["variable costs/CU"])
                                           + ((hydrogen.loc["ID_wasserstoff_storage"]["periodical costs/CU"] + hydrogen.loc["ID_wasserstoff_storage"]["variable costs/CU"])
                                           + hydrogen_energycosts_gesamt)) / hydrogen_output_kWh
    hydrogengehstehungskosten_kg =  hydrogengehstehungskosten_kWh * 33.333
    


#Wasserstoffbrennstoffzelle
brennstoffzelle_kw = components.loc["ID_pemfc-brennstoffzelle_transformer"]["investment/kW"]
brennstoffzelle_invest = components.loc["ID_pemfc-brennstoffzelle_transformer"]["periodical costs/CU"] * 20 + components.loc["ID_pemfc-brennstoffzelle_transformer"]["variable costs/CU"] * 20
brennstoffzelle_energycosts = components.loc["ID_pemfc-brennstoffzelle_transformer"]["constraints/CU"]
brennstoffzelle_input = components.loc["ID_pemfc-brennstoffzelle_transformer"]["input 1/kWh"]
brennstoffzelle_output = components.loc["ID_pemfc-brennstoffzelle_transformer"]["output 1/kWh"]
if brennstoffzelle_kw == 0:
    anteil_brennstoffzelle = 0.0
    brennstoffzelle_stromgehstehungskosten = 0.0
else:
    anteil_brennstoffzelle = brennstoffzelle_input / hydrogen_output_kWh
    brennstoffzelle_stromgehstehungskosten = (brennstoffzelle_invest/20 + brennstoffzelle_energycosts + brennstoffzelle_input * hydrogengehstehungskosten_kWh ) / brennstoffzelle_output 
    

#E-Mobilität
e_ladesaeule_kw = components.loc["ID_e-ladesaeule_link"]["investment/kW"]
e_ladesaeule_output = components.loc["ID_e-ladesaeule_link"]["output 1/kWh"]
e_ladesaeule_invest = components.loc["ID_e-ladesaeule_link"]["periodical costs/CU"] * 20 + components.loc["ID_e-ladesaeule_link"]["variable costs/CU"] * 20
e_lkw_kw = e_lkw["investment/kW"].sum()
e_lkw_invest = e_lkw["periodical costs/CU"].sum() * 20 + e_lkw["variable costs/CU"].sum() * 20
e_lkw_energycosts = e_lkw["constraints/CU"].sum()
e_lkw_km = e_lkw["output 1/kWh"].sum()
wind_in_elektromobilität_summe = a["wind_in_elektromobilität"].sum()
pv_in_elektromobilität_summe = a["pv_in_elektromobilität"].sum()
netz_in_elektromobilität_summe = a["netz_in_elektromobilität"].sum()


elektromobilität_energycosts_strombezug = pv_in_elektromobilität_summe * einspeiseverguetung_pv + wind_in_H2_summe * einspeiseverguetung_wind + netz_in_elektromobilität_summe * strombezugspreis
if e_lkw_kw  == 0:
    elektromobilität_km_gehstehungskosten = 0.0
    Erneuerbar_in_elektromobilität_anteil = 0.0
else:
    elektromobilität_km_gehstehungskosten = (e_lkw_invest/20 + e_ladesaeule_invest/20 + e_lkw_energycosts + elektromobilität_energycosts_strombezug ) / e_lkw_km 
    Erneuerbar_in_elektromobilität_anteil = (wind_in_elektromobilität_summe + pv_in_elektromobilität_summe) / a["(('ID_niederspannung_bus', 'ID_e-ladesaeule_link'), 'flow')"].sum()


#H2-Mobilität
wasserstofftankstelle_kw = components.loc["ID_wasserstofftankstelle_link"]["investment/kW"]
wasserstofftankstelle_invest = components.loc["ID_wasserstofftankstelle_link"]["periodical costs/CU"] * 20 + components.loc["ID_wasserstofftankstelle_link"]["variable costs/CU"] * 20
wasserstoff_lkw_kw = wasserstoff_lkw["investment/kW"].sum()
wasserstoff_lkw_invest = wasserstoff_lkw["periodical costs/CU"].sum() * 20 + wasserstoff_lkw["variable costs/CU"].sum() * 20
wasserstoff_lkw_energycosts = wasserstoff_lkw["constraints/CU"].sum()
wasserstoff_lkw_km = wasserstoff_lkw["output 1/kWh"].sum()

wasserstofftankstelle_input = components.loc["ID_wasserstofftankstelle_link"]["input 1/kWh"]
wasserstoff_mobilität_energycosts_bezug = wasserstofftankstelle_input * hydrogengehstehungskosten_kWh

if wasserstofftankstelle_kw == 0:
    wasserstoff_mobilität_km_gehstehungskosten = 0.0
else:
    wasserstoff_mobilität_km_gehstehungskosten = (wasserstoff_lkw_invest/20 + wasserstofftankstelle_invest/20 + wasserstoff_lkw_energycosts + wasserstoff_mobilität_energycosts_bezug) / wasserstoff_lkw_km

#Diesel-Mobilität
diesel_lkw_kw = diesel_lkw["investment/kW"].sum()
diesel_lkw_invest = diesel_lkw["periodical costs/CU"].sum() * 20 + diesel_lkw["variable costs/CU"].sum() * 20
diesel_lkw_energycosts = diesel_lkw["constraints/CU"].sum()
diesel_lkw_km = diesel_lkw["output 1/kWh"].sum()

diesel_bezug = components.loc["ID_diesel-tankstelle_link"]["input 1/kWh"]
energycosts_diesel_bezug = diesel_bezug * dieselbezugspreis

if diesel_lkw_kw == 0:
    diesel_mobilität_km_gehstehungskosten = 0.0
else:
    diesel_mobilität_km_gehstehungskosten = (diesel_lkw_invest/20 + diesel_lkw_energycosts + energycosts_diesel_bezug) / diesel_lkw_km


#Mobilität
gesamt_km = diesel_lkw_km + wasserstoff_lkw_km + e_lkw_km
anteil_diesel = diesel_lkw_km *100 / gesamt_km 
anteil_wasserstoff = wasserstoff_lkw_km *100 / gesamt_km
anteil_elektro = e_lkw_km *100 / gesamt_km

anteil_erneuerbar_wasserstoff = Erneuerbar_in_H2_anteil * anteil_wasserstoff
anteil_erneuerbar_elektro = Erneuerbar_in_elektromobilität_anteil * anteil_elektro

autarkie_mobilität = anteil_erneuerbar_elektro + anteil_erneuerbar_wasserstoff



#CO2-Emission in g
diesel_emission = components.loc["ID_diesel_bus_shortage"]["output 1/kWh"] * 266 
strom_emission = components.loc["ID_netzanschluss_bus_shortage"]["output 1/kWh"] * 28.8
einsparung_einspeisung = (pv_einspeisung * -61 + wind_einspeisung * -9)


pv_emission = pv_energy_output * 61
wind_emission = wind_energy_output * 9
elektrolyse_emission = hydrogen_output_kWh * 1.3 
brennstoffzelle_emission = brennstoffzelle_output * 10
hydrogen_storage_emission = hydrogen_storage_kWh * 320
battery_emission = battery_kwh * 3960
e_ladesaeule_emission = 0
diesel_lkw_vehicle_cycle_emission = diesel_lkw_km * 83 * 1.609 
wasserstoff_lkw_vehicle_cycle_emission = wasserstoff_lkw_km * 100 * 1.609
e_lkw_vehicle_cycle_emission = e_lkw_km * 240 * 1.609

emission_in_g = (diesel_emission + strom_emission + e_ladesaeule_emission
                 + pv_emission + wind_emission + elektrolyse_emission + brennstoffzelle_emission + hydrogen_storage_emission + battery_emission
                 + diesel_lkw_vehicle_cycle_emission + wasserstoff_lkw_vehicle_cycle_emission + e_lkw_vehicle_cycle_emission - einsparung_einspeisung)
emission_in_t = emission_in_g / 1000000


#Stromverbrauch
netzbezug = components.loc["ID_netzanschluss_bus_shortage"]["output 1/kWh"]
beschaffung_gesamt = netzbezug + wind_eigenverbrauch + pv_eigenverbrauch + battery_energy_output + brennstoffzelle_output
anteil_wind = wind_eigenverbrauch *100 / beschaffung_gesamt
anteil_pv = pv_eigenverbrauch *100 / beschaffung_gesamt
anteil_battery =  battery_energy_output *100 / beschaffung_gesamt
anteil_brennstoffzelle = brennstoffzelle_output *100 / beschaffung_gesamt
anteil_netz =  netzbezug *100 / beschaffung_gesamt
autarkie_strom = 100 - anteil_netz


#Ausgabe
datensatz = pd.DataFrame([[invest_mehr, energycosts_gesamt, energycosts_reduction, gewinn_nach_20_jahren,amortisation, rentabilität,
                           beschaffung_gesamt, autarkie_strom, anteil_netz, anteil_pv, anteil_wind, anteil_battery, anteil_brennstoffzelle,
                           gesamt_km, autarkie_mobilität, anteil_diesel, anteil_elektro, anteil_wasserstoff, emission_in_t,
                           pv_kw, pv_invest, pv_energycosts, pv_energy_output, pv_eigenverbrauch_prozentual, pv_stromgehstehungskosten,
                           wind_kw, wind_invest, wind_energycosts, wind_energy_output, wind_eigenverbrauch_prozentual, wind_stromgehstehungskosten,
                           erneuerbar_invest, erneuerbar_energycosts,erneuerbar_energy_output, erneuerbar_eigenverbrauch_prozentual, erneuerbar_stromgehstehungskosten,
                           battery_kwh, battery_invest, battery_energycosts, battery_energy_output, battery_stromgehstehungskosten,umschlagszahl,netzentgelt,leistung,
                           elektrolyse_kw, hydrogen_storage_kWh, hydrogen_invest, hydrogen_energycosts_gesamt, hydrogengehstehungskosten_kg,
                           brennstoffzelle_kw, brennstoffzelle_invest, brennstoffzelle_energycosts, brennstoffzelle_output, brennstoffzelle_stromgehstehungskosten,
                           diesel_lkw_invest, diesel_lkw_energycosts, diesel_lkw_km, diesel_mobilität_km_gehstehungskosten,
                           e_lkw_invest, e_lkw_energycosts, e_lkw_km, elektromobilität_km_gehstehungskosten, e_ladesaeule_kw, e_ladesaeule_invest,
                           wasserstoff_lkw_invest, wasserstoff_lkw_energycosts, wasserstoff_lkw_km, wasserstoff_mobilität_km_gehstehungskosten, wasserstofftankstelle_kw, wasserstofftankstelle_invest
                           ]], columns=["Mehrinvestitionskosten in €", "Laufende Energiekosten in €/a", "Kostenreduktion in €/a", "Gewinn nach 20 Jahren in €","Amortisationszeit in Jahren", "Rentabilität in %",
                                        "Energieverbrauch Strom in kWh", "Autarkiegrad Strom in %" ,
                                        "Deckungsanteil: Netz in %", "Deckungsanteil: PV in %", "Deckungsanteil: Wind in %", "Deckungsanteil: Batteriespeicher in %", "Deckungsanteil: Brenstoffzelle in %",
                                        "LKW: Gefahrene Kilometer", "Autarkiegrad Mobilität in %" ,"Anteil Diesel-Mobilität  in %", "Anteil Elektro-Mobilität in %", "Anteil Wasserstoff-Mobilität in %","CO2-Emissioinen in t/a",
                                        "PV: Installierte Leistung in kWp", "PV: Investitionskosten in €", "PV: Laufende Kosten in €/a",
                                        "PV: Energieoutput in kWh", "PV: Eigenverbrauch in %", "PV: Stromgehstehungskosten in €/kWh",
                                        "Wind: Installierte Leistung in kW", "Wind: Investitionskosten in €", "Wind: Laufende Kosten in €/a",
                                        "Wind: Energieoutput in kWh", "Wind: Eigenverbrauch in %", "Wind: Stromgehstehungskosten €/kWh",
                                        "Erneuerbare Energien: Investitionskosten in €", "Erneuerbare Energien: Laufende Kosten in €/a",
                                        "Erneuerbare Energien: Energieoutput in kWh", "Erneuerbare Energien: Eigenverbrauch in %", "Erneuerbare Energien: Stromgehstehungskosten €/kWh",
                                        "Batterie: Speichergröße in kWh", "Batterie: Investitionskosten in €", "Batterie: Laufende Kosten in €/a",
                                        "Batterie: Energieoutput in kWh", "Batterie: Stromgehstehungskosten in €/kWh","Batterie: Umschlagshäufigkeit","Netzentgelt in €/a",  "Bezugsleistung in kVA",
                                        "Elektrolyse-Leistung in kW", "Wasserstoffspeicher in kWh", "Wasserstoffproduktion: Investitionskosten in €",
                                        "Wasserstoffproduktion: Laufende Kosten in €/a", "Wasserstoffgestehungskosten €/kg",
                                        "Brennstoffzelle: Installierte Leistung in kW", "Brennstoffzelle: Investitionskosten in €", "Brennstoffzelle: Laufende Kosten in €/a",
                                        "Brennstoffzelle: Energieoutput in kWh", "Brennstoffzelle: Stromgehstehungskosten in €/kWh",
                                        "Diesel-Lkw: Investitionskosten in €", "Diesel-Lkw: Laufende Kosten in €/a", "Diesel-Lkw: Gefahrene Kilometer", "Diesel-Lkw: Kosten je km",
                                        "E-Lkw: Investitionskosten in €", "E-Lkw: Laufende Kosten in €/a", "E-Lkw: Gefahrene Kilometer", "E-Lkw: Kosten je km",
                                        "E-Ladesäule: Installierte Leistung in kW", "E-Ladesäule: Investitionskosten in €",
                                        "Wasserstoff-Lkw: Investitionskosten in €", "Wasserstoff-Lkw: Laufende Kosten in €/a", "Wasserstoff-Lkw: Gefahrene Kilometer", "Wasserstoff-Lkw: Kosten je km",
                                        "Wasserstofftankstelle: Installierte Leistung in kW", "Wasserstofftankstelle: Investitionskosten in €"])

ergebnisstabelle = pd.concat([ergebnisstabelle] + [datensatz], ignore_index = True)


print(ergebnisstabelle)
ergebnisstabelle.to_excel("Ergebnisse.xlsx")
