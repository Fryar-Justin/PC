package com.inductiveautomation.apachepoi;

import com.inductiveautomation.ignition.client.gateway_interface.GatewayConnectionManager;
import com.inductiveautomation.ignition.common.Dataset;

import java.io.Serializable;
import java.util.Date;

/**
 * Created by SKMUser on 5/4/2016.
 */
public class ExcelFunctions {

    public static final String MODULE_ID = "com.inductiveautomation.apachepoi.reporting-apache-poi-template";

    public static byte[] getParkMeadowsSheet(Date reportDate, Dataset turbidityData, Dataset Hours, Dataset WQData) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getParkMeadowsSheet", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, turbidityData, Hours, WQData});
    }

    public static byte[] getQuinnsMonitoring(Date reportDate, Dataset productionData, Dataset rackResults, Dataset chemicals) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getQuinnsMonitoring", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, productionData, rackResults, chemicals});
    }

    public static byte[] getJSSD(Date reportDate, Dataset fiveMinData, Dataset turbData, Dataset hours) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getJSSD", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, fiveMinData, turbData, hours});
    }

    public static byte[] getCreekside(Date reportDate, Dataset fiveMinData, Dataset turbData, Dataset hours) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getCreekside", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, fiveMinData, turbData, hours});
    }

    public static byte[] getCreeksideUVDaily(Date reportDate, Dataset runHours, Dataset totalProd, Dataset redData, Dataset offSpecData) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getCreeksideUVDaily", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, runHours, totalProd, redData, offSpecData});
    }

    public static byte[] getCreeksideUVMonthly(Date reportDate, Dataset runHours, Dataset totalProd, Dataset offSpecData) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getCreeksideUVMonthly", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, runHours, totalProd, offSpecData});
    }

    public static byte[] getCreeksideUVOffSpec(Date reportDate, Dataset offSpecData) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getCreeksideUVOffSpec", new Class[]{Date.class, Dataset.class}, new Serializable[]{reportDate, offSpecData});
    }

    public static byte[] getQuinnsSheet(Date reportDate, Dataset FiveMinData, Dataset rackResults, Dataset WQData) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getQuinnsSheet", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, FiveMinData, rackResults, WQData});
    }

    public static byte[] getQuinnsSheetMnO2(Date reportDate, Dataset FiveMinData, Dataset rackResults, Dataset WQData) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getQuinnsSheetMnO2", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, FiveMinData, rackResults, WQData});
    }

    public static byte[] getQuinnsFlows(Date reportDate, Dataset sewerFlows) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getQuinnsFlows", new Class[]{Date.class, Dataset.class}, new Serializable[]{reportDate, sewerFlows});
    }


    public static byte[] getOgdensSheet(Date reportDate, Dataset FiveMinData, Dataset rackResults, Dataset WQData, Dataset turbidity, Dataset rack2Results, Dataset rack3Results, Dataset rack4Results,
                                        Dataset rack5Results, Dataset rack6Results, Dataset rack7Results, Dataset rack8Results, Dataset rack9Results) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getOgdensSheet", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class, Dataset.class,
                Dataset.class, Dataset.class, Dataset.class, Dataset.class, Dataset.class, Dataset.class, Dataset.class, Dataset.class},
                new Serializable[]{reportDate, FiveMinData, rackResults, WQData, turbidity, rack2Results, rack3Results, rack4Results, rack5Results, rack6Results, rack7Results, rack8Results, rack9Results});
    }

    public static byte[] getGroundWaterDisinfection(Date reportDate, Dataset groundWaterData, Dataset hypoSpeed) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getGroundWaterDisinfection", new Class[]{Date.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, groundWaterData, hypoSpeed});
    }

    public static byte[] getGroundWaterDisinfectionNoPM(Date reportDate, Dataset groundWaterData, Dataset hypoSpeed) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getGroundWaterDisinfectionNoPM", new Class[]{Date.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, groundWaterData, hypoSpeed});
    }

    public static byte[] getMembraneReport(Date reportDate, Dataset productionData, Dataset rackResults, Dataset IT_data) throws Exception {
        return GatewayConnectionManager.getInstance().getGatewayInterface().moduleInvokeSafe(MODULE_ID, "getMembraneReport", new Class[]{Date.class, Dataset.class, Dataset.class, Dataset.class}, new Serializable[]{reportDate, productionData, rackResults, IT_data});
    }
}
