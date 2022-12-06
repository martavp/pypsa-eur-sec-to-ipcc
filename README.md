# PyPSA-Eur-Sec to IPCC/IAMC

This repository contains scripts to export IAMC format datafiles from PyPSA-Eur-Sec networks. 

Formatting follows the requirements for the IPCC AR6 database that can be found here: [IPCC AR6 documentation](https://data.ene.iiasa.ac.at/ar6-scenario-submission/#/about)

[Information on the IAMC databeses](https://software.ene.iiasa.ac.at/ixmp-server/tutorials.html)

## Exporting a dataset

Using the [pypsa_to_IPCC.py](pypsa_to_IPCC.py) script a scenario run with the PyPSA-Eur-Sec model can be exportet to a dataset in IAMC format. In the top of the script information regarting the scenario must be updated.

~~~
model = "PyPSA-Eur-Sec 0.0.2" 
model_version = "0.0.6"
literature_reference = "Pedersen, T. T., Gøtske, E. K., Dvorak, A., Andresen, G. B., & Victoria, M. (2022). Long-term implications of reduced gas imports on the decarbonization of the European energy system. Joule, 6(7), 1566-1580."
climate_target = 21 # CO2 budget

scenarios={'Base_1.5':'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb25.7ex0_',
            'Gaslimit_1.5':'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb25.7ex0-gasconstrained_',
            'Base_2':'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb73.9ex0_',
            'Gaslimit_2': 'postnetworks/elec_s370_37m_lv1.0__3H-T-H-B-I-solar+p3-dist1-cb73.9ex0-gasconstrained_',
                     }
~~~

Running the script will result in a single .xlsx file per scenario located in the results folder. 