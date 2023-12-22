# -*- coding: utf-8 -*-
"""
Created on Tue Jul 21 15:58:59 2020

@author: 20174250
"""
#Based on Dutch GRID
def RunPP4():
    #-----------Import modules-----
    import pandas as pd
    import pandapower as pp
    import pandapower.plotting as plot
    import matplotlib.pyplot as plt
    #import seaborn
    from pandapower.plotting import simple_plot, simple_plotly, pf_res_plotly
    import pandapower.topology
    
    #-----------Test results-----
    #import pandapower.test
    #pandapower.test.run_all_tests()
    
    #-----------Start calculation-----
    net = pp.create_empty_network()
    #pd.set_option('display.max_columns', 30)
    
    #Set bus
    df = pd.read_excel("bus1.xlsx", sheet_name="bus", index_col=0)
    for idx in df.index:
        pp.create_bus(net, *df.loc[idx,:])
        #pp.create_bus(net, vn_kv=df.at[idx, "vn_kv"])  
    #print(net.bus)
    
    # df = pd.read_excel("generator1.xlsx", sheet_name="generator", index_col=0)
    # for idx in df.index:
    #     pp.create_gen(net, *df.loc[idx,:])
    # #print(net.gen)    

    #set Loads
    df = pd.read_excel("load1.xlsx", sheet_name="load", index_col=0)
    for idx in df.index:
        pp.create_load(net, *df.loc[idx,:])# 
        #pp.create_load(net, bus=df.at[idx, "bus"], p_mw=df.at[idx, "p_mw"])  
    #print(net.load)

    #Set exit grid
    df = pd.read_excel("ext_grid1.xlsx", sheet_name="ext_grid", index_col=0)
    for idx in df.index:
        pp.create_ext_grid(net, *df.loc[idx,:]) 
        #pp.create_ext_grid(net, bus=df.at[idx, "bus"], vm_pu=df.at[idx, "vm_pu"], va_degree=df.at[idx, "va_degree"])  
    #print(net.ext_grid)

    #Set line
    df = pd.read_excel("line1.xlsx", sheet_name="line", index_col=0)
    for idx in df.index:
        pp.create_line_from_parameters(net, *df.loc[idx,:])#, g_us_per_km=df.at[idx, "g_us_per_km"], max_i_ka=df.at[idx, "max_i_ka"])       
    #print(net.line)
    
    #----------Execute power flow calculation and results
    pp.runpp(net, calculate_voltage_angles=True, init="dc") #High voltage network can be considered as DC network. Only voltage angles considerd.
    #pp.runpp(net, calculate_voltage_angles=True, init="dc") #https://github.com/e2nIEE/pandapower/blob/master/tutorials/powerflow.ipynb
    #print(net.res_bus)
    #print(net.res_line)
    OutTableBus = net.res_bus
    OutTableLine = net.res_line
    #print(OutTableLine)
    #tb.lf_info(net)
    #OutTableBus.to_csv(r'C:\Users\20174250\OneDrive - TU Eindhoven\Master thesis JvdB\Energy system model\Zenmo-ZERO-Netherlands-JvdB-June-2020_V1.1\outputBus.csv', index = False)
    #OutTableLine.to_csv(r'C:\Users\20174250\OneDrive - TU Eindhoven\Master thesis JvdB\Energy system model\Zenmo-ZERO-Netherlands-JvdB-June-2020_V1.1\outputLine.csv', index = False)
    #OutTableBus.to_excel(r'C:\Users\20174250\Documents\PandaPowerOutput\outputBusTest.xlsx', index = False)
    #OutTableLine.to_excel(r'C:\Users\20174250\Documents\PandaPowerOutput\outputLineTest.xlsx', index = False)

    #-----------Set variables according to Anylogic variables
    _380KvCrayestein_Maasvlakte_zwart_ = OutTableLine.iloc[0,13]
    _380KvGeertruidenberg_Rilland_zwart_MADE = OutTableLine.iloc[1,13]
    _380KvRilland_Borssele_zwart_MADE = OutTableLine.iloc[2,13]
    _380KvLelystad_Diemen_zwart_ = OutTableLine.iloc[3,13]
    _380KvDodewaard_Doetinchem_zwart_ = OutTableLine.iloc[4,13]
    _380KvDodewaard_Boxmeer_zwart_MADE = OutTableLine.iloc[5,13]
    _380KvDoetinchem_Hengelo_zwart_ = OutTableLine.iloc[6,13]
    _380KvEemshaven_Meeden_zwart_ = OutTableLine.iloc[7,13]
    _380KvZwolle_ens_zwart_ = OutTableLine.iloc[8,13]
    _380KvEindhoven_geertruidenberg_zwart_ = OutTableLine.iloc[9,13]
    _380KvKrimpen_Bleiswijk_zwart_ = OutTableLine.iloc[10,13]
    _380KvKrimpen_Geertruidenberg_zwart_ = OutTableLine.iloc[11,13]
    _380KvEns_lelystad_zwart_ = OutTableLine.iloc[12,13]
    _380KvMaasbracht_Boxmeer_zwart_MADE = OutTableLine.iloc[13,13]
    _380KvMaasbracht_Eindhoven_zwart_ = OutTableLine.iloc[14,13]
    _380KvOostzaan_Beverwijk_zwart_ = OutTableLine.iloc[15,13]
    _380KvHengelo_Zwolle_zwart_ = OutTableLine.iloc[16,13]
    _380KvZwolle_meeden_zwart_ = OutTableLine.iloc[17,13]
    _380KvDiemen_oostzaan = OutTableLine.iloc[18,13]
    _380KvOostzaan_Beverwijk_zwart_ = OutTableLine.iloc[19,13]
    _380KvKrimpen_Crayestein_zwart_ = OutTableLine.iloc[20,13]
    offshoreBorssele = OutTableLine.iloc[21,13]
    offshoreIJmuiden = OutTableLine.iloc[22,13]
    offshoreHoekvHolland = OutTableLine.iloc[23,13]
    ex_importBelgium1 = OutTableLine.iloc[24,13]
    ex_importBelgium2 = OutTableLine.iloc[25,13]
    ex_importGermany1 = OutTableLine.iloc[26,13]
    ex_importGermany2 = OutTableLine.iloc[27,13]
    ex_importGermany3 = OutTableLine.iloc[28,13]
    ex_importScandavian = OutTableLine.iloc[29,13]
    ex_importUK = OutTableLine.iloc[30,13]
    _380KvMaasvlakte_Blieswijk_MADE = OutTableLine.iloc[31,13]
    _380KvKrimpen_Breukelen_MADE = OutTableLine.iloc[32,13]
    _380KvBreukelen_Diemen_MADE = OutTableLine.iloc[33,13]
    
    #pp.plotting.create_bus_collection(net)
    
    #print(net.line)
    #print(net.bus)
    #print(net.ext_grid)
    #pp.plotting.simple_plot(net)
    #simple_plotly(net)
    #pf_res_plotly(net)
    #pandapower.topology.create_nxgraph(net)
    #pp.diagnostic(net)
    return _380KvCrayestein_Maasvlakte_zwart_, _380KvGeertruidenberg_Rilland_zwart_MADE, _380KvRilland_Borssele_zwart_MADE, _380KvLelystad_Diemen_zwart_, _380KvDodewaard_Doetinchem_zwart_, _380KvDodewaard_Boxmeer_zwart_MADE, _380KvDoetinchem_Hengelo_zwart_, _380KvEemshaven_Meeden_zwart_, _380KvZwolle_ens_zwart_, _380KvEindhoven_geertruidenberg_zwart_, _380KvKrimpen_Bleiswijk_zwart_, _380KvKrimpen_Geertruidenberg_zwart_, _380KvEns_lelystad_zwart_, _380KvMaasbracht_Boxmeer_zwart_MADE, _380KvMaasbracht_Eindhoven_zwart_, _380KvOostzaan_Beverwijk_zwart_, _380KvHengelo_Zwolle_zwart_, _380KvZwolle_meeden_zwart_, _380KvDiemen_oostzaan, _380KvOostzaan_Beverwijk_zwart_, _380KvKrimpen_Crayestein_zwart_, offshoreBorssele, offshoreIJmuiden, offshoreHoekvHolland, ex_importBelgium1, ex_importBelgium2, ex_importGermany1, ex_importGermany2, ex_importGermany3, ex_importScandavian, ex_importUK, _380KvMaasvlakte_Blieswijk_MADE, _380KvKrimpen_Breukelen_MADE, _380KvBreukelen_Diemen_MADE

#Based on own modified excel files
def RunPP6():
    import pandas as pd
    import pandapower as pp
    import pandapower.plotting as plot
    import matplotlib.pyplot as plt
    import seaborn
    from pandapower.plotting import simple_plot, simple_plotly, pf_res_plotly
    import pandapower.topology
    
    net = pp.create_empty_network()
    #pd.set_option('display.max_columns', 30)
    
    #Set bus
    df = pd.read_excel("bus.xlsx", sheet_name="bus", index_col=0)
    for idx in df.index:
        pp.create_bus(net, *df.loc[idx,:])
        #pp.create_bus(net, vn_kv=df.at[idx, "vn_kv"])  
    #print(net.bus)

    #set Loads
    df = pd.read_excel("load.xlsx", sheet_name="load", index_col=0)
    for idx in df.index:
        pp.create_load(net, *df.loc[idx,:])# 
        #pp.create_load(net, bus=df.at[idx, "bus"], p_mw=df.at[idx, "p_mw"])  
    #print(net.load)

    #Set exit grid
    df = pd.read_excel("ext_grid.xlsx", sheet_name="ext_grid", index_col=0)
    for idx in df.index:
        pp.create_ext_grid(net, *df.loc[idx,:]) 
        #pp.create_ext_grid(net, bus=df.at[idx, "bus"], vm_pu=df.at[idx, "vm_pu"], va_degree=df.at[idx, "va_degree"])  
    #print(net.ext_grid)

    #Set line
    df = pd.read_excel("line.xlsx", sheet_name="line", index_col=0)
    for idx in df.index:
        pp.create_line_from_parameters(net, *df.loc[idx,:])#, g_us_per_km=df.at[idx, "g_us_per_km"], max_i_ka=df.at[idx, "max_i_ka"])       
    #print(net.line)

    #Set Trafo
    df = pd.read_excel("trafo.xlsx", sheet_name="trafo", index_col=0)
    for idx in df.index:
        pp.create_transformer(net, *df.loc[idx,:])      
        #pp.create_transformer(net, hv_bus=df.at[idx, "hv_bus"], lv_bus=df.at[idx, "lv_bus"], std_type=df.at[idx, "std_type"],)
    
    pp.runpp(net)
    print(net.res_bus)
    print(net.res_line)
    OutTableBus = net.res_bus
    OutTableLine = net.res_line
    #tb.lf_info(net)
    #OutTableBus.to_csv(r'C:\Users\20174250\OneDrive - TU Eindhoven\Master thesis JvdB\Energy system model\Zenmo-ZERO-Netherlands-JvdB-June-2020_V1.1\outputBus.csv', index = False)
    #OutTableLine.to_csv(r'C:\Users\20174250\OneDrive - TU Eindhoven\Master thesis JvdB\Energy system model\Zenmo-ZERO-Netherlands-JvdB-June-2020_V1.1\outputLine.csv', index = False)
    OutTableBus.to_excel(r'C:\Users\20174250\Documents\PandaPowerOutput\outputBusTest.xlsx', index = False)
    OutTableLine.to_excel(r'C:\Users\20174250\Documents\PandaPowerOutput\outputLineTest.xlsx', index = False)
    line0_1 = OutTableLine.iloc[0,13]
    line1_2 = OutTableLine.iloc[1,13]
    #print(line0_1)
    #print(line1_2)
    #print(net.line)
    #print(net.bus)
    print(net.ext_grid)
    #pp.plotting.simple_plot(net)
    #simple_plotly(net)
    #pf_res_plotly(net)
    #pandapower.topology.create_nxgraph(net)
    
    return line0_1, line1_2

#Based on Youtube video
def RunPP5():
    import pandas as pd
    import pandapower as pp

    import pandapower.plotting as plot
    import matplotlib.pyplot as plt
    import seaborn
    from pandapower.plotting import simple_plot, simple_plotly, pf_res_plotly
    import pandapower.topology
    net = pp.create_empty_network()
    #d.set_option('display.max_columns', 30)
    
    #Set bus
    df = pd.read_excel("NetworkHVPandaPower2.xlsx", sheet_name="bus", index_col=0)
    for idx in df.index:
        pp.create_bus(net, *df.loc[idx,:])
        #pp.create_bus(net, vn_kv=df.at[idx, "vn_kv"])  
    print(net.bus)

    #set Loads
    df = pd.read_excel("NetworkHVPandaPower2.xlsx", sheet_name="load", index_col=0)
    for idx in df.index:
        pp.create_load(net, *df.loc[idx,:])# 
        #pp.create_load(net, bus=df.at[idx, "bus"], p_mw=df.at[idx, "p_mw"])  
    print(net.load)

    #Set exit grid
    df = pd.read_excel("NetworkHVPandaPower2.xlsx", sheet_name="ext_grid", index_col=0)
    for idx in df.index:
        pp.create_ext_grid(net, *df.loc[idx,:]) 
        pp.create_ext_grid(net, bus=df.at[idx, "bus"], vm_pu=df.at[idx, "vm_pu"], va_degree=df.at[idx, "va_degree"])  
    print(net.ext_grid)

    #Set line
    df = pd.read_excel("NetworkHVPandaPower2.xlsx", sheet_name="line", index_col=0)
    for idx in df.index:
        pp.create_line_from_parameters(net, *df.loc[idx,:])#, g_us_per_km=df.at[idx, "g_us_per_km"], max_i_ka=df.at[idx, "max_i_ka"])       
    print(net.line)

    #Set Trafo
    df = pd.read_excel("NetworkHVPandaPower2.xlsx", sheet_name="trafo", index_col=0)
    for idx in df.index:
        pp.create_transformer(net, *df.loc[idx,:])      
        #pp.create_transformer(net, hv_bus=df.at[idx, "hv_bus"], lv_bus=df.at[idx, "lv_bus"], std_type=df.at[idx, "std_type"],)
    pp.runpp(net)
    print(net.res_bus)
    print(net.res_line)
    OutTableBus = net.res_bus
    OutTableLine = net.res_line
    #tb.lf_info(net)
    OutTableBus.to_csv(r'C:\Users\20174250\OneDrive - TU Eindhoven\Master thesis JvdB\Energy system model\Zenmo-ZERO-Netherlands-JvdB-June-2020_V1.1\outputBus.csv', index = False)
    OutTableLine.to_csv(r'C:\Users\20174250\OneDrive - TU Eindhoven\Master thesis JvdB\Energy system model\Zenmo-ZERO-Netherlands-JvdB-June-2020_V1.1\outputLine.csv', index = False)
    
    #pp.plotting.simple_plot(net)
    #pf_res_plotly(net)
    #pandapower.topology.create_nxgraph(net)
    return