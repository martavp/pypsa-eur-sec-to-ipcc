# -*- coding: utf-8 -*-
"""
Script to convert networks from PyPSA-Eur-Sec v0.0.2 to data format used in the
IPCC AR6 database
"""

import pypsa
import openpyxl
import pandas as pd

#original IPCC file, official template
template_path = "global_sectoral/IPCC_AR6_WG3_Global_sectoral_Pathways_scenario_template_v3.1_20221027.xlsx"
#use the official template

# Metadata
model = "PyPSA-Eur-Sec 0.0.2"
model_version = "0.0.6"
literature_reference = "Pedersen, T. T., Gøtske, E. K., Dvorak, A., Andresen, G. B., & Victoria, M. (2022). Long-term implications of reduced gas imports on the decarbonization of the European energy system. Joule, 6(7), 1566-1580."
climate_target = 21 # CO2 budget

wind_split = ['DE', 'ES', 'FI', 'FR', 'GB', 'IT', 'NO', 'PL', 'RO', 'SE']

scenarios={'Base_1.5':'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb25.7ex0_',
            'Gaslimit_1.5':'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb25.7ex0-gasconstrained_',
            'Base_2':'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb73.9ex0_',
            'Gaslimit_2': 'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb73.9ex0-gasconstrained_',
                     }


output_folder = 'results/'

years = [2020, 2025, 2030, 2035, 2040, 2045, 2050]

countries = ['AT','BE','BG','CH','CZ','DE','DK','EE','ES','FI','FR','GB','GR','HR',
             'HU','IT','LT','LU','LV','NO','PL','PT','RO','SE','SI','SK','IE', 'NL',
             #'RS','BA'
             ] 

iso2name={'AT':'Austria',
          'BE':'Belgium',
          'BG':'Bulgaria',
          'CH':'Switzerland',
          'CZ':'Czech Republic',
          'DE':'Germany',
          'DK':'Denmark',
          'EE':'Estonia',
          'ES':'Spain',
          'FI':'Finland',
          'FR':'France',
          'GB':'United Kingdom',
          'GR':'Greece',
          'HR':'Croatia',
          'HU':'Hungary',
          'IT':'Italy',
          'LT':'Lithuania',
          'LU':'Luxembourg',
          'LV':'Latvia',
          'NO':'Norway',
          'PL':'Poland',
          'PT':'Portugal',
          'RO':'Romania',
          'SE':'Sweden',
          'SI':'Slovenia',
          'SK':'Slovakia',
          'IE':'Ireland',
          'NL':'The Netherlands',}

for scenario in scenarios:
    #one excel file per scenario
    file = openpyxl.load_workbook(template_path)
    ds = file['data'] #data sheet
    for year in years:
        n = pypsa.Network(f"{scenarios[scenario]}{year}.nc")
        costs = pd.read_csv(f"costs/costs_{year}.csv", index_col=[0,1])
        
        col=[c for c in ds[1] if c.value==year][0].column
        
        for i,country in enumerate(countries):
            if year == 2020:
                #one datasheet per country including information from different years
                target = file.copy_worksheet(file['data'])
                target.title ='data' + str(i)
            ds = file['data' + str(i)] 
            var={}
            
            """
            Capacity : Solar PV, onshore and offshore wind
            """
            #MW -> GW
            var['Capacity|Electricity|Solar|PV'] =0.001*n.generators.p_nom_opt[country + ' solar'] if country + ' solar' in n.generators.index else 0
            var['Capacity|Electricity|Solar'] = var['Capacity|Electricity|Solar|PV']
            var['Capacity |Electricity|Solar|PV|Rooftop PV'] = 0.5 * var['Capacity|Electricity|Solar|PV']
            var['Capacity|Electricity|Solar|PV|Utility-scale PV'] = 0.5 * var['Capacity|Electricity|Solar|PV']
            if country in wind_split:
                var['Capacity|Electricity|Wind|Onshore']=0.001*n.generators.p_nom_opt[[i for i in n.generators.index if country in i and 'onwind' in i]].sum() 
                var['Capacity|Electricity|Wind|Offshore']=0.001*n.generators.p_nom_opt[[i for i in n.generators.index if country in i and 'offwind' in i]].sum() 
            else:
                var['Capacity|Electricity|Wind|Onshore']=0.001*n.generators.p_nom_opt[country + ' onwind'] if country + ' onwind' in n.generators.index else 0
                var['Capacity|Electricity|Wind|Offshore']=0.001*n.generators.p_nom_opt[country + ' offwind'] if country + ' offwind' in n.generators.index else 0
            var['Capacity|Electricity|Wind']=var['Capacity|Electricity|Wind|Onshore']+var['Capacity|Electricity|Wind|Offshore']


            """
            Capacity : Nuclear, Coal, Lignite, OCGT, CCGT, Biomass
            """
            #MW -> GW
            var['Capacity|Electricity|Nuclear'] =0.001*n.links.efficiency[country + ' nuclear']*n.links.p_nom_opt[country + ' nuclear'] if country + ' nuclear' in n.links.index else 0
            var['Capacity|Electricity|Coal|w/o CCS'] = 0.001*n.links.efficiency[country + ' coal']*n.links.p_nom_opt[country + ' coal'] if country + ' coal' in n.links.index else 0
            var['Capacity|Electricity|Coal|w/o CCS'] += 0.001*n.links.efficiency[country + ' lignite']*n.links.p_nom_opt[country + ' lignite'] if country + ' lignite' in n.links.index else 0
            var['Capacity|Electricity|Coal'] =var['Capacity|Electricity|Coal|w/o CCS'] 
            var['Capacity|Electricity|Gas|w/o CCS'] = 0.001* n.links.efficiency[country + ' OCGT']*n.links.p_nom_opt[country + ' OCGT'] if country + ' OCGT' in n.links.index else 0
            var['Capacity|Electricity|Gas|w/o CCS'] += 0.001*n.links.efficiency[country + ' CCGT']*n.links.p_nom_opt[country + ' CCGT'] if country + ' CCGT' in n.links.index else 0
            var['Capacity|Electricity|Gas'] = var['Capacity|Electricity|Gas|w/o CCS']
            var['Capacity|Electricity|Biomass|w/o CCS'] = 0.001* n.links.efficiency[country + ' biomass EOP']*n.links.p_nom_opt[country + ' biomass EOP'] if country + ' biomass EOP' in n.links.index else 0
            var['Capacity|Electricity|Biomass|w/o CCS'] += 0.001* n.links.efficiency[country + ' central biomass CHP electric']*n.links.p_nom_opt[country + ' central biomass CHP electric'] if country + ' central biomass CHP electric' in n.links.index else 0
            var['Capacity|Electricity|Biomass'] = var['Capacity|Electricity|Biomass|w/o CCS']
            
            """
            Capacity : hydro (reservoir, ror)
            """
            #MW -> GW
            var['Capacity|Electricity|Hydro'] = 0.001*n.storage_units.p_nom_opt[country + ' hydro'] if country + ' hydro' in n.storage_units.index else 0
            var['Capacity|Electricity|Hydro'] += 0.001*n.generators.p_nom_opt[country + ' ror'] if country + ' ror' in n.generators.index else 0
            
            """
            Electricity : Solar PV, onshore and offshore wind
            """
            #MWh -> EJ
            var['Secondary Energy|Electricity|Solar|PV'] = 3.6e-9*n.generators_t.p[country + ' solar'].sum() if country + ' solar' in n.generators.index else 0
            var['Secondary Energy|Electricity|Solar'] = var['Secondary Energy|Electricity|Solar|PV']
            var['Secondary Energy|Electricity|Solar|PV|Rooftop PV'] = 0.5 * var['Secondary Energy|Electricity|Solar|PV'] 
            var['Secondary Energy|Electricity|Solar|PV|Utility-scale PV'] = 0.5 * var['Secondary Energy|Electricity|Solar|PV'] 
            
            if country in wind_split:
                var['Secondary Energy|Electricity|Wind|Onshore'] = 3.6e-9*n.generators_t.p[[i for i in n.generators.index if country in i and 'onwind' in i]].sum().sum()
                var['Secondary Energy|Electricity|Wind|Offshore'] = 3.6e-9*n.generators_t.p[[i for i in n.generators.index if country in i and 'offwind' in i]].sum().sum()
            else:
                var['Secondary Energy|Electricity|Wind|Onshore'] = 3.6e-9*n.generators_t.p[country + ' onwind'].sum() if country + ' onwind' in n.generators.index else 0
                var['Secondary Energy|Electricity|Wind|Offshore'] = 3.6e-9*n.generators_t.p[country + ' offwind'].sum() if country + ' offwind' in n.generators.index else 0
            var['Secondary Energy|Electricity|Wind'] = var['Secondary Energy|Electricity|Wind|Onshore'] + var['Secondary Energy|Electricity|Wind|Offshore']
            
            """
            Electricity : Nuclear, Coal, Lignite, OCGT, CCGT, biomass
            """
            #MWh -> EJ
            var['Secondary Energy|Electricity|Nuclear'] = -3.6e-9*n.links_t.p1[country + ' nuclear'].sum() if country + ' nuclear' in n.links.index else 0
            var['Secondary Energy|Electricity|Coal|w/o CCS'] =- 3.6e-9*n.links_t.p1[country + ' coal'].sum() if country + ' coal' in n.links.index else 0
            var['Secondary Energy|Electricity|Coal|w/o CCS'] += -3.6e-9*n.links_t.p1[country + ' lignite'].sum() if country + ' lignite' in n.links.index else 0
            var['Secondary Energy|Electricity|Gas|w/o CCS'] = -3.6e-9*n.links_t.p1[country + ' OCGT'].sum() if country + ' OCGT' in n.links.index else 0
            var['Secondary Energy|Electricity|Gas|w/o CCS'] += -3.6e-9*n.links_t.p1[country + ' CCGT'].sum() if country + ' CCGT' in n.links.index else 0
            var['Secondary Energy|Electricity|Biomass|w/o CCS'] = -3.6e-9*n.links_t.p1[country + ' biomass EOP'].sum() if country + ' biomass EOP' in n.links.index else 0
            var['Secondary Energy|Electricity|Biomass|w/o CCS'] += -3.6e-9*n.links_t.p1[country + ' central biomass CHP electric'].sum() if country + ' central biomass CHP electric' in n.links.index else 0
            
            """
            Electricity : Hydro (reservoir, ror)
            """
            #MWh -> EJ
            var['Secondary Energy|Electricity|Hydro'] = 3.6e-9*n.storage_units_t.p[country + ' hydro'].sum() if country + ' hydro' in n.storage_units.index else 0
            var['Secondary Energy|Electricity|Hydro'] += 3.6e-9*n.generators_t.p[country + ' ror'].sum() if country + ' ror' in n.generators.index else 0

            """
            Capacity : storage (PHS, battery, H2 storage)
            """
            #MWh to GWh
            var['Capacity|Electricity|Storage|Pumped Hydro Storage'] = 0.001 *n.storage_units.p_nom_opt[country + ' PHS'] if country + ' PHS' in n.storage_units.index else 0
            var['Capacity|Electricity|Storage|Battery Capacity|Utility-scale Battery'] = 0.001 *n.stores.e_nom_opt[country + ' battery'] if country + ' battery' in n.stores.index else 0
            var['Capacity|Electricity|Storage|Battery Capacity'] = var['Capacity|Electricity|Storage|Battery Capacity|Utility-scale Battery']
            #variable Hydrogen Storage Capacity not included in EU Climate Advisory Board Scenario Explorer
            #var['Capacity|Electricity|Storage|Hydrogen Storage Capacity|overground'] = 0.001 *n.stores.e_nom_opt[country + ' H2 Store tank'] if country + ' H2 Store tank' in n.stores.index else 0
            #var['Capacity|Electricity|Storage|Hydrogen Storage Capacity|underground'] = 0.001 *n.stores.e_nom_opt[country + ' H2 Store underground'] if country + ' H2 Store underground' in n.stores.index else 0
            #var['Capacity|Electricity|Storage|Hydrogen Storage Capacity'] = (var['Capacity|Electricity|Storage|Hydrogen Storage Capacity|overground']
            #                                                                + var['Capacity|Electricity|Storage|Hydrogen Storage Capacity|underground'])
            var['Capacity|Electricity|Storage|Hydrogen Storage Capacity'] = (0.001 *n.stores.e_nom_opt[country + ' H2 Store tank'] if country + ' H2 Store tank' in n.stores.index else 0
            + 0.001 *n.stores.e_nom_opt[country + ' H2 Store underground'] if country + ' H2 Store underground' in n.stores.index else 0)
            var['Capacity|Electricity|Storage Capacity'] = ( var['Capacity|Electricity|Storage|Pumped Hydro Storage']
                                                            +  var['Capacity|Electricity|Storage|Battery Capacity']
                                                            + var['Capacity|Electricity|Storage|Hydrogen Storage Capacity'])
            
            """
            Capacity : heat pumps, heat resistors, Sabatier (synthetic gas)
            """
            # ELectric capacity
            # MW to Gw
            var['Capacity|Heating|Heat pumps'] = 0.001*n.links.p_nom_opt[country + ' central heat pump'] if country + ' central heat pump' in n.links.index else 0
            var['Capacity|Heating|Heat pumps'] += 0.001*n.links.p_nom_opt[country + ' decentral heat pump'] if country + ' decentral heat pump' in n.links.index else 0
            var['Capacity|Heating|Electric boilers'] = 0.001*n.links.p_nom_opt[country + ' central resistive heater'] if country + ' central resistive heater' in n.links.index else 0
            var['Capacity|Heating|Electric boilers'] += 0.001*n.links.p_nom_opt[country + ' decentral resistive heater'] if country + ' decentral resistive heater' in n.links.index else 0
            var['Capacity|Gas|Synthetic'] = 0.001*n.links.p_nom_opt[country + ' Sabatier'] if country + ' Sabatier' in n.links.index else 0
            
            """
            Final Energy (heating) : heat pumps, heat resistors, Sabatier (synthetic gas)
            """
            #MWh -> EJ
            #50/50 services/domestic
            var['Final Energy|Residential and Commercial|Commercial|Heating|Heat pumps'] = - 3.6e-9 * 0.5 * n.links_t.p1[country + ' central heat pump'].sum() if country + ' central heat pump' in n.links.index else 0
            var['Final Energy|Residential and Commercial|Commercial|Heating|Heat pumps'] += - 3.6e-9 * 0.5 * n.links_t.p1[country + ' decentral heat pump'].sum() if country + ' decentral heat pump' in n.links.index else 0
            var['Final Energy|Residential and Commercial|Residential|Heating|Heat pumps'] = - 3.6e-9 * 0.5 * n.links_t.p1[country + ' central heat pump'].sum() if country + ' central heat pump' in n.links.index else 0
            var['Final Energy|Residential and Commercial|Residential|Heating|Heat pumps'] += - 3.6e-9  *0.5 * n.links_t.p1[country + ' decentral heat pump'].sum() if country + ' decentral heat pump' in n.links.index else 0
            var['Final Energy|Residential and Commercial|Commercial|Heating|Electric boilers'] = - 3.6e-9 * 0.5 * n.links_t.p1[country + ' central resistive heater'].sum() if country + ' central resistive heater' in n.links.index else 0
            var['Final Energy|Residential and Commercial|Commercial|Heating|Electric boilers'] += - 3.6e-9 * 0.5 * n.links_t.p1[country + ' decentral resistive heater'].sum() if country + ' decentral resistive heater' in n.links.index else 0
            var['Final Energy|Residential and Commercial|Residential|Heating|Electric boilers'] = - 3.6e-9 * 0.5 * n.links_t.p1[country + ' central resistive heater'].sum() if country + ' central resistive heater' in n.links.index else 0
            var['Final Energy|Residential and Commercial|Residential|Heating|Electric boilers'] += - 3.6e-9 * 0.5 * n.links_t.p1[country + ' decentral resistive heater'].sum() if country + ' decentral resistive heater' in n.links.index else 0
            var['Final Energy|Gas|Synthetic'] = - 3.6e-9*n.links_t.p1[country + ' Sabatier'].sum() if country + ' Sabatier' in n.links.index else 0
            
            """
            Capacity : Electrolysis
            """
            #MW to GW
            var['Capacity|Hydrogen|Electricity'] = 0.001*n.links.p_nom_opt[country + ' H2 Electrolysis']*n.links.efficiency[country + ' H2 Electrolysis'] if country + ' H2 Electrolysis' in n.links.index else 0

            """
            Hydrogen production
            """
            #MWh to EJ
            var['Secondary Energy|Hydrogen|Electricity'] = -3.6e-9*n.links_t.p1[country + ' H2 Electrolysis'].sum() if country + ' H2 Electrolysis' in n.links.index else 0
            
            """
            Capital cost and Lifetime 
            """
            # €_2015/kW  to US$_2010/kW
            EUR2015_USD2010 = 1.11 /1.09
            ipcc2pypsa ={'Capital Cost': 'investment',
                         'Lifetime':'lifetime'} 
            for metric in ['Capital Cost', 'Lifetime']:
                factor = EUR2015_USD2010 if metric =='Capital Cost' else 1
                var[metric + ' |Electricity|Solar|PV|Rooftop PV'] = factor * costs.loc[('solar-rooftop', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Solar|PV|Utility-scale PV'] = factor * costs.loc[('solar-utility', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Solar|PV'] = 0.5* var[metric + ' |Electricity|Solar|PV|Rooftop PV'] + 0.5*var[metric + '|Electricity|Solar|PV|Utility-scale PV']
                var[metric + '|Electricity|Wind|Onshore'] = factor * costs.loc[('onwind', ipcc2pypsa[metric]), 'value']
                var[metric + '|Electricity|Wind|Offshore'] = factor * costs.loc[('offwind', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Nuclear'] = factor * costs.loc[('nuclear', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Coal|w/o CCS'] = factor * costs.loc[('coal', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Gas|w/o CCS'] = factor * costs.loc[('OCGT', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Hydro'] = factor * costs.loc[('hydro', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Storage|Pumped Hydro Storage'] = factor * costs.loc[('PHS', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Storage|Battery Capacity|Utility-scale Battery'] = factor * costs.loc[('battery storage', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Storage|Battery Capacity'] = var [metric + '|Electricity|Storage|Battery Capacity|Utility-scale Battery']
                var[metric + '|Heating|Heat pumps'] = factor * costs.loc[('decentral air-sourced heat pump', ipcc2pypsa[metric]),'value']
                var[metric + '|Heating|Electric boilers'] = factor * costs.loc[('decentral resistive heater', ipcc2pypsa[metric]),'value']
                var[metric + '|Gas|Synthetic'] = factor * costs.loc[('methanation', ipcc2pypsa[metric]),'value']
                var[metric + '|Storage|Thermal Energy Storage|Household storage'] = factor * costs.loc[('decentral water tank storage', ipcc2pypsa[metric]),'value']
                var[metric + '|Storage|Thermal Energy Storage|District heating storage'] = factor * costs.loc[('central water tank storage', ipcc2pypsa[metric]),'value']
                var[metric + '|Hydrogen|Electricity'] = factor * costs.loc[('electrolysis', ipcc2pypsa[metric]),'value']
                var[metric + '|Electricity|Biomass|w/o CCS'] = factor * costs.loc[('biomass EOP', ipcc2pypsa[metric]),'value']
            
            #there is a spelling error with variables overground/underground with capital/small intitial
            # the following lines can be included in the loop when the spelling mistake is corrected
            metric='Capital Cost'
            #var[metric + '|Electricity|Storage|Hydrogen Storage Capacity|overground'] = factor * costs.loc[('hydrogen storage tank', ipcc2pypsa[metric]),'value']
            #var[metric + '|Electricity|Storage|Hydrogen Storage Capacity|underground'] = factor * costs.loc[('hydrogen storage underground', ipcc2pypsa[metric]),'value']
            
            metric='Lifetime'
            #var[metric + '|Electricity|Storage|Hydrogen Storage Capacity|Overground'] = factor * costs.loc[('hydrogen storage tank', ipcc2pypsa[metric]),'value']
            #var[metric + '|Electricity|Storage|Hydrogen Storage Capacity|Underground'] = factor * costs.loc[('hydrogen storage underground', ipcc2pypsa[metric]),'value']
            
            """
            OM Cost
            """
            #US$_2010/kW·year
            factor = EUR2015_USD2010
            var['OM Cost |Electricity|Solar|PV|Rooftop PV'] = factor * 0.01*costs.loc[('solar-rooftop', 'FOM'),'value']*costs.loc[('solar-rooftop', 'investment'),'value']
            var['OM Cost|Electricity|Solar|PV|Utility-scale PV'] = factor * 0.01*costs.loc[('solar-utility', 'FOM'),'value']*costs.loc[('solar-utility', 'investment'),'value']
            var['OM Cost|Fixed|Electricity|Solar|PV'] = 0.5* var['OM Cost |Electricity|Solar|PV|Rooftop PV'] + 0.5*var['OM Cost|Electricity|Solar|PV|Utility-scale PV']
            var['OM Cost|Fixed|Electricity|Wind|Onshore'] = factor * 0.01*costs.loc[('onwind', 'FOM'),'value']*costs.loc[('onwind', 'investment'),'value']
            var['OM Cost|Fixed|Electricity|Wind|Offshore'] = factor * 0.01*costs.loc[('offwind', 'FOM'),'value']*costs.loc[('offwind', 'investment'),'value']
            var['OM Cost|Fixed|Electricity|Nuclear'] = factor * 0.01*costs.loc[('nuclear', 'FOM'),'value']*costs.loc[('nuclear', 'investment'),'value']
            var['OM Cost|Fixed|Electricity|Coal|w/o CCS'] = factor * 0.01*costs.loc[('coal', 'FOM'),'value']*costs.loc[('coal', 'investment'),'value']
            var['OM Cost|Fixed|Electricity|Gas|w/o CCS'] = factor * 0.01*costs.loc[('OCGT', 'FOM'),'value']*costs.loc[('OCGT', 'investment'),'value']
            var['OM Cost|Fixed|Electricity|Hydro'] = factor * 0.01*costs.loc[('hydro', 'FOM'),'value']*costs.loc[('hydro', 'investment'),'value']
            var['OM Cost|Electricity|Storage|Pumped Hydro Storage'] = factor * 0.01*costs.loc[('PHS', 'FOM'),'value']*costs.loc[('PHS', 'investment'),'value']
            var['OM Cost|Electricity|Storage|Battery Capacity|Utility-scale Battery'] = factor * 0.01*costs.loc[('battery storage', 'FOM'),'value']*costs.loc[('battery storage', 'investment'),'value']
            var['OM Cost|Electricity|Storage|Battery Capacity'] = var [metric + '|Electricity|Storage|Battery Capacity|Utility-scale Battery']
            #var['OM Cost|Electricity|Storage|Hydrogen Storage Capacity|Overground'] = factor * 0.01*costs.loc[('hydrogen storage tank', 'FOM'),'value']*costs.loc[('hydrogen storage tank', 'investment'),'value']
            #var['OMCost|Electricity|Storage|Hydrogen Storage Capacity|Underground'] = factor * 0.01*costs.loc[('hydrogen storage underground', 'FOM'),'value']*costs.loc[('hydrogen storage underground', 'investment'),'value']
            var['OM Cost|Heating|Heat pumps'] =factor * 0.01*costs.loc[('decentral air-sourced heat pump', 'FOM'),'value']*costs.loc[('decentral air-sourced heat pump', 'investment'),'value']
            var['OM Cost|Heating|Electric boilers'] = factor * 0.01*costs.loc[('decentral resistive heater', 'FOM'),'value']*costs.loc[('decentral resistive heater', 'investment'),'value']
            var['OM Cost|Gas|Synthetic'] = factor * 0.01*costs.loc[('methanation', 'FOM'),'value']*costs.loc[('methanation', 'investment'),'value']
            var['OM Cost|Storage|Thermal Energy Storage|Household storage'] = factor * 0.01*costs.loc[('decentral water tank storage', 'FOM'),'value']*costs.loc[('decentral water tank storage', 'investment'),'value']
            var['OM Cost|Storage|Thermal Energy Storage|District heating storage'] = factor * 0.01*costs.loc[('central water tank storage', 'FOM'),'value']*costs.loc[('central water tank storage', 'investment'),'value']
            var['OM Cost|Fixed|Hydrogen|Electricity'] = factor * 0.01*costs.loc[('electrolysis', 'FOM'),'value']*costs.loc[('electrolysis', 'investment'),'value']
            var['OM Cost|Fixed|Electricity|Biomass|w/o CCS'] = factor * 0.01*costs.loc[('biomass EOP', 'FOM'),'value']*costs.loc[('biomass EOP', 'investment'),'value']

            """
            Efficiency
            """
            var['Effciency|Heating|Heat pumps'] = costs.loc[('decentral air-sourced heat pump', 'efficiency'),'value']
            var['Efficiency|Heating|Electric boilers'] = costs.loc[('decentral resistive heater', 'efficiency'),'value']
            var['Efficiency|Gas|Synthetic'] = costs.loc[('methanation', 'efficiency'),'value']
            var['Efficiency|Hydrogen|Electricity'] = costs.loc[('electrolysis', 'efficiency'),'value']
            var['Efficiency|Electricity|Biomass|w/o CCS'] = costs.loc[('biomass EOP', 'efficiency'),'value']


            for v in var.keys():
                ro=[r for r in ds['D'] if r.value==v][0].row
                ds.cell(row=ro, column=col).value = round(var[v],3) 
                ds.cell(row=ro, column=1).value = model
                ds.cell(row=ro, column=2).value = scenario
                ds.cell(row=ro, column=3).value = iso2name[country] #region
    # add scenario name to 'meta_scenario' sheet
    ds2 = file['meta_scenario']
    ds2.cell(row=4, column=2).value = scenario
    ds2.cell(row=4, column=4).value = model
    ds2.cell(row=4, column=5).value = model_version
    ds2.cell(row=4, column=10).value = literature_reference
    ds2.cell(row=4, column=14).value = climate_target

 
        
    file.save(f"{output_folder}/IPCC_AR6_{scenario}.xlsx")

