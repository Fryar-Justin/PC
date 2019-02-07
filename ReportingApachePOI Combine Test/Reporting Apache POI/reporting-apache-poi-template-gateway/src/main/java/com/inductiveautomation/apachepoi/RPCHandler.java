package com.inductiveautomation.apachepoi;

import com.inductiveautomation.ignition.common.Dataset;

import java.util.Date;

/**
 * Created by SKMUser on 5/4/2016.
 */
public class RPCHandler {

    public RPCHandler(){

    }
    public byte[] getQuinnsFlows(Date reportDate, Dataset sewerFlows) throws Exception {
        return ExcelFunctions.getQuinnsFlows(reportDate, sewerFlows);
    }
    public byte[] getQuinnsMonitoring(Date reportDate, Dataset productionData, Dataset rackResults, Dataset chemicals) throws Exception {
        return ExcelFunctions.getParkMeadowsSheet(reportDate, productionData, rackResults, chemicals);
    }
    public byte[] getParkMeadowsSheet(Date reportDate, Dataset turbidityData, Dataset Hours, Dataset WQData) throws Exception {
        return ExcelFunctions.getParkMeadowsSheet(reportDate, turbidityData, Hours, WQData);
    }
    public byte[] getQuinnsSheet(Date reportDate, Dataset turbidityData, Dataset Hours, Dataset WQData) throws Exception {
        return ExcelFunctions.getQuinnsSheet(reportDate, turbidityData, Hours, WQData);
    }

    public byte[] getCreekside(Date reportDate, Dataset fiveMinData, Dataset turbData, Dataset hours) throws Exception {
        return ExcelFunctions.getCreekside(reportDate, fiveMinData, turbData, hours);
    }

    public byte[] getJSSD(Date reportDate, Dataset fiveMinData, Dataset turbData, Dataset hours) throws Exception {
        return ExcelFunctions.getJSSD(reportDate, fiveMinData, turbData, hours);
    }

    public byte[] getCreeksideUVDaily(Date reportDate, Dataset runHours, Dataset totalProd, Dataset redData, Dataset offSpecData) throws Exception {
        return ExcelFunctions.getCreeksideUVDaily(reportDate, runHours, totalProd, redData, offSpecData);
    }

    public byte[] getCreeksideUVMonthly(Date reportDate, Dataset runHours, Dataset totalProd, Dataset offSpecData) throws Exception {
        return ExcelFunctions.getCreeksideUVMonthly(reportDate, runHours, totalProd, offSpecData);
    }

    public byte[] getCreeksideUVOffSpec(Date reportDate, Dataset offSpecData) throws Exception {
        return ExcelFunctions.getCreeksideUVOffSpec(reportDate, offSpecData);
    }


    public byte[] getQuinnsSheetMnO2(Date reportDate, Dataset turbidityData, Dataset Hours, Dataset WQData) throws Exception {
        return ExcelFunctions.getQuinnsSheetMnO2(reportDate, turbidityData, Hours, WQData);
    }

    public byte[] getOgdensSheet(Date reportDate, Dataset FiveMinData, Dataset turbidity, Dataset rackResults, Dataset WQData, Dataset rack2Results, Dataset rack3Results,
                                 Dataset rack4Results, Dataset rack5Results, Dataset rack62Results, Dataset rack7Results, Dataset rack8Results, Dataset rack9Results) throws Exception {
        return ExcelFunctions.getOgdensSheet(reportDate, FiveMinData, turbidity, rackResults, WQData, rack2Results, rack3Results, rack4Results, rack5Results, rack62Results, rack7Results,
                rack8Results, rack9Results);
    }
    public byte[] getGroundWaterDisinfection(Date reportDate, Dataset groundWaterData, Dataset hypoSpeed) throws Exception {
        return ExcelFunctions.getGroundWaterDisinfection(reportDate, groundWaterData, hypoSpeed);
    }

    public byte[] getGroundWaterDisinfectionNoPM(Date reportDate, Dataset groundWaterData, Dataset hypoSpeed) throws Exception {
        return ExcelFunctions.getGroundWaterDisinfectionNoPM(reportDate, groundWaterData, hypoSpeed);
    }

    public byte[] getMembraneReport(Date reportDate, Dataset productionData, Dataset rackResults, Dataset IT_data) throws Exception {
        return ExcelFunctions.getMembraneReport(reportDate, productionData, rackResults, IT_data);
    }
}
