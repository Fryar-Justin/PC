package com.inductiveautomation.apachepoi;

import com.inductiveautomation.ignition.common.Dataset;
import com.inductiveautomation.ignition.gateway.sqltags.scanclasses.MultiDriverExecutableScanClass;
import com.inductiveautomation.reporting.common.api.QueryResults;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import java.util.HashMap;
import java.util.Map;

import static org.apache.poi.poifs.crypt.HashAlgorithm.none;

/**
 * Created by Travis Cox on 11/12/2015.
 * Highly edited by Allen Rogers 4/29/16 to add Park City Reports
 * Edited for the Quinn's Reports 7/20/2016
 */
public class ExcelFunctions {

    static final Logger logger = LoggerFactory.getLogger("POI Functions");

    static SimpleDateFormat displayFormat = new SimpleDateFormat("M/d/yyyy");
    static SimpleDateFormat excelDateFormat = new SimpleDateFormat("M/d/yy h:mm a");
    static SimpleDateFormat operationalDateFormat = new SimpleDateFormat("M/d/yyyy h:mm a");
    static SimpleDateFormat excelShortDateFormat = new SimpleDateFormat("M/d");
    static SimpleDateFormat monthFormat = new SimpleDateFormat("MMMM");
    static SimpleDateFormat hoursMinsSecs = new SimpleDateFormat("HH:mm:ss");

    //--------------------------------------------Method to handle Datasets and QueryResults
    //Park City's Method Declarations
    public static byte[] getParkMeadowsSheet(Date reportDate, Dataset turbidityData, Dataset Hours, Dataset WQData) throws Exception {
        return _getParkMeadowsSheet(reportDate, new ObjectDatasetWrapper(turbidityData), new ObjectDatasetWrapper(Hours), new ObjectDatasetWrapper(WQData));
    }

    public static byte[] getParkMeadowsSheet(Date reportDate, QueryResults turbidityData, QueryResults Hours, QueryResults WQData) throws Exception {
        return _getParkMeadowsSheet(reportDate, new ObjectDatasetWrapper(turbidityData), new ObjectDatasetWrapper(Hours), new ObjectDatasetWrapper(WQData));
    }

    public static byte[] getCreekside(Date reportDate, Dataset fiveMinData, Dataset turbData, Dataset hours) throws Exception {
        return _getCreekside(reportDate, new ObjectDatasetWrapper(fiveMinData), new ObjectDatasetWrapper(turbData), new ObjectDatasetWrapper(hours));
    }

    public static byte[] getCreekside(Date reportDate, QueryResults fiveMinData, QueryResults turbData, QueryResults hours) throws Exception {
        return _getCreekside(reportDate, new ObjectDatasetWrapper(fiveMinData), new ObjectDatasetWrapper(turbData), new ObjectDatasetWrapper(hours));
    }

    public static byte[] getJSSD(Date reportDate, Dataset fiveMinData, Dataset turbData, Dataset hours) throws Exception {
        return _getJSSD(reportDate, new ObjectDatasetWrapper(fiveMinData), new ObjectDatasetWrapper(turbData), new ObjectDatasetWrapper(hours));
    }

    public static byte[] getJSSD(Date reportDate, QueryResults fiveMinData, QueryResults turbData, QueryResults hours) throws Exception {
        return _getJSSD(reportDate, new ObjectDatasetWrapper(fiveMinData), new ObjectDatasetWrapper(turbData), new ObjectDatasetWrapper(hours));
    }

    public static byte[] getCreeksideUVDaily(Date reportDate, Dataset runHours, Dataset totalProd, Dataset redData, Dataset offSpecData) throws Exception {
        return _getCreeksideUVDaily(reportDate, new ObjectDatasetWrapper(runHours), new ObjectDatasetWrapper(totalProd), new ObjectDatasetWrapper(redData), new ObjectDatasetWrapper(offSpecData));
    }

    public static byte[] getCreeksideUVDaily(Date reportDate, Dataset runHours, QueryResults totalProd, QueryResults redData, Dataset offSpecData) throws Exception {
        return _getCreeksideUVDaily(reportDate, new ObjectDatasetWrapper(runHours), new ObjectDatasetWrapper(totalProd), new ObjectDatasetWrapper(redData), new ObjectDatasetWrapper(offSpecData));
    }

    public static byte[] getCreeksideUVMonthly(Date reportDate, Dataset runHours, Dataset totalProd, Dataset offSpecData) throws Exception {
        return _getCreeksideUVMonthly(reportDate, new ObjectDatasetWrapper(runHours), new ObjectDatasetWrapper(totalProd), new ObjectDatasetWrapper(offSpecData));
    }

    public static byte[] getCreeksideUVMonthly(Date reportDate, Dataset runHours, QueryResults totalProd, QueryResults offSpecData) throws Exception {
        return _getCreeksideUVMonthly(reportDate, new ObjectDatasetWrapper(runHours), new ObjectDatasetWrapper(totalProd), new ObjectDatasetWrapper(offSpecData));
    }

    public static byte[] getCreeksideUVOffSpec(Date reportDate, Dataset offSpecData) throws Exception {
        return _getCreeksideUVOffspec(reportDate, new ObjectDatasetWrapper(offSpecData));
    }

    public static byte[] getCreeksideUVOffSpec(Date reportDate, QueryResults offSpecData) throws Exception {
        return _getCreeksideUVOffspec(reportDate, new ObjectDatasetWrapper(offSpecData));
    }

    //Quinns Method Declarations

    public static byte[] getQuinnsSheetMnO2(Date reportDate, Dataset FiveMinData, Dataset rackResults, Dataset WQData) throws Exception {
        return _getQuinnsSheetMnO2(reportDate, new ObjectDatasetWrapper(FiveMinData), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(WQData));
    }

    public static byte[] getQuinnsSheetMnO2(Date reportDate, QueryResults FiveMinData, QueryResults rackResults, QueryResults WQData) throws Exception {
        return _getQuinnsSheetMnO2(reportDate, new ObjectDatasetWrapper(FiveMinData), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(WQData));
    }


    public static byte[] getQuinnsSheet(Date reportDate, Dataset FiveMinData, Dataset rackResults, Dataset WQData) throws Exception {
        return _getQuinnsSheet(reportDate, new ObjectDatasetWrapper(FiveMinData), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(WQData));
    }

    public static byte[] getGroundWaterDisinfection(Date reportDate, Dataset groundWaterData, Dataset hypoSpeed) throws Exception {
        return _getGroundWaterDisinfection(reportDate, new ObjectDatasetWrapper(groundWaterData),  new ObjectDatasetWrapper(hypoSpeed));
    }

    public static byte[] getGroundWaterDisinfection(Date reportDate, QueryResults groundWaterData, QueryResults hypoSpeed) throws Exception {
        return _getGroundWaterDisinfection(reportDate, new ObjectDatasetWrapper(groundWaterData), new ObjectDatasetWrapper(hypoSpeed));
    }

    public static byte[] getGroundWaterDisinfectionNoPM(Date reportDate, Dataset groundWaterData, Dataset hypoSpeed) throws Exception {
        return _getGroundWaterDisinfectionNoPM(reportDate, new ObjectDatasetWrapper(groundWaterData),  new ObjectDatasetWrapper(hypoSpeed));
    }

    public static byte[] getGroundWaterDisinfectionNoPM(Date reportDate, QueryResults groundWaterData, QueryResults hypoSpeed) throws Exception {
        return _getGroundWaterDisinfectionNoPM(reportDate, new ObjectDatasetWrapper(groundWaterData), new ObjectDatasetWrapper(hypoSpeed));
    }

    public static byte[] getQuinnsSheet(Date reportDate, QueryResults FiveMinData, QueryResults rackResults, QueryResults WQData) throws Exception {
        return _getQuinnsSheet(reportDate, new ObjectDatasetWrapper(FiveMinData), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(WQData));
    }

    public static byte[] getQuinnsFlows(Date reportDate, Dataset SewerFlows) throws Exception {
        return _getQuinnsFlows(reportDate, new ObjectDatasetWrapper(SewerFlows));
    }

    public static byte[] getQuinnsFlows(Date reportDate, QueryResults SewerFlows) throws Exception {
        return _getQuinnsFlows(reportDate, new ObjectDatasetWrapper(SewerFlows));
    }

    public static byte[] getOgdensSheet(Date reportDate, Dataset FiveMinData, Dataset rackResults, Dataset WQData, Dataset turbidity,
                                        Dataset rack2Results, Dataset rack3Results, Dataset rack4Results, Dataset rack5Results, Dataset rack6Results,
                                        Dataset rack7Results, Dataset rack8Results, Dataset rack9Results) throws Exception {
        return _getOgdensSheet(reportDate, new ObjectDatasetWrapper(FiveMinData), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(WQData), new ObjectDatasetWrapper(turbidity),
                new ObjectDatasetWrapper(rack2Results), new ObjectDatasetWrapper(rack3Results), new ObjectDatasetWrapper(rack4Results), new ObjectDatasetWrapper(rack5Results),
                new ObjectDatasetWrapper(rack6Results), new ObjectDatasetWrapper(rack7Results), new ObjectDatasetWrapper(rack8Results), new ObjectDatasetWrapper(rack9Results));
    }

    public static byte[] getOgdensSheet(Date reportDate, QueryResults FiveMinData, QueryResults rackResults, QueryResults WQData, QueryResults turbidity,
                                        QueryResults rack2Results, QueryResults rack3Results, QueryResults rack4Results, QueryResults rack5Results, QueryResults rack6Results,
                                        QueryResults rack7Results, QueryResults rack8Results, QueryResults rack9Results) throws Exception {
        return _getOgdensSheet(reportDate, new ObjectDatasetWrapper(FiveMinData), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(WQData), new ObjectDatasetWrapper(turbidity),
                new ObjectDatasetWrapper(rack2Results), new ObjectDatasetWrapper(rack3Results), new ObjectDatasetWrapper(rack4Results),  new ObjectDatasetWrapper(rack5Results),
                new ObjectDatasetWrapper(rack6Results),  new ObjectDatasetWrapper(rack7Results),  new ObjectDatasetWrapper(rack8Results),  new ObjectDatasetWrapper(rack9Results));
    }

    //Quinns Monitoring data
    public static byte[] getQuinnsMonitoring(Date reportDate, Dataset production, Dataset rackResults, Dataset chemicals) throws Exception {
        return _getQuinnsMonitoringData(reportDate, new ObjectDatasetWrapper(production), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(chemicals));
    }
    public static byte[] getQuinnsMonitoring(Date reportDate, QueryResults production, QueryResults rackResults, Dataset chemicals) throws Exception {
        return _getQuinnsMonitoringData(reportDate, new ObjectDatasetWrapper(production), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(chemicals));
    }

    // New
    public static byte[] getMembraneReport(Date reportDate, Dataset production, Dataset rackResults, Dataset IT_data) throws Exception {
        return _getMembraneReport(reportDate, new ObjectDatasetWrapper(production), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(IT_data));
    }
    public static byte[] getMembraneReport(Date reportDate, QueryResults production, QueryResults rackResults, QueryResults IT_data) throws Exception {
        return _getMembraneReport(reportDate, new ObjectDatasetWrapper(production), new ObjectDatasetWrapper(rackResults), new ObjectDatasetWrapper(IT_data));
    }
    //----------------------------------------------Functions-----------------------------------------------------------


    private static byte[] _getParkMeadowsSheet(Date reportDate, ObjectDatasetWrapper turbidityData, ObjectDatasetWrapper Hours, ObjectDatasetWrapper WQData) throws Exception {
        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);

        InputStream is = ExcelFunctions.class.getResourceAsStream("Park Meadows Well.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);
        int numDaysInMonth = cal.getActualMaximum(Calendar.DAY_OF_MONTH);

        XSSFSheet sheet = wb.getSheet("Turb Compliance");
        sheet.getRow(4).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(5).getCell(2).setCellValue(year);

        // Turbidity Worksheet
        sheet = wb.getSheet("Turb Data");
        cal.setTime(reportDate);
        sheet.getRow(2).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(3).getCell(2).setCellValue(year);

        Double blank = sheet.getRow(11).getCell(7).getNumericCellValue();

        // Clear all data first
        for (int i = 0; i < 31; i++) {
            Integer rowIdx = 11 + i;
            sheet.getRow(rowIdx).getCell(0).setCellValue("");
            sheet.getRow(rowIdx).getCell(1).setCellValue(0);
        }

        for (ObjectDatasetWrapper.Row row : turbidityData) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double fourAMTurb = (Double) row.getKeyValue("PRKM_4AM_TURB");
            Double eightAMTurb = (Double) row.getKeyValue("PRKM_8AM_TURB");
            Double twelvePMTurb = (Double) row.getKeyValue("PRKM_12PM_TURB");
            Double fourPMTurb = (Double) row.getKeyValue("PRKM_4PM_TURB");
            Double eightPMTurb = (Double) row.getKeyValue("PRKM_8PM_TURB");
            Double twelveAMTurb = (Double) row.getKeyValue("PRKM_12AM_TURB");

            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer hour = cal.get(Calendar.HOUR_OF_DAY);
            Integer rowIdx = 11 + (day - 1);

            sheet.getRow(rowIdx).getCell(0).setCellValue(excelShortDateFormat.format(t_stamp));
            if(fourAMTurb == null)
            {
                sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
            }
            else
            {
                sheet.getRow(rowIdx).getCell(2).setCellValue(fourAMTurb);
            }
            if(eightAMTurb == null)
            {
                sheet.getRow(rowIdx).getCell(3).setCellValue("PO");
            }
            else
            {
                sheet.getRow(rowIdx).getCell(3).setCellValue(eightAMTurb);
            }
            if(twelvePMTurb == null)
            {
                sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
            }
            else
            {
                sheet.getRow(rowIdx).getCell(4).setCellValue(twelvePMTurb);
            }
            if(fourPMTurb == null)
            {
                sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
            }
            else
            {
                sheet.getRow(rowIdx).getCell(5).setCellValue(fourPMTurb);
            }
            if(eightPMTurb == null)
            {
                sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
            }
            else
            {
                sheet.getRow(rowIdx).getCell(6).setCellValue(eightPMTurb);
            }
            if(twelveAMTurb == null)
            {
                sheet.getRow(rowIdx).getCell(7).setCellValue("PO");
            }
            else
            {
                sheet.getRow(rowIdx).getCell(7).setCellValue(twelveAMTurb);
            }

        }

        for (ObjectDatasetWrapper.Row row : Hours) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");

            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = 11 + (day - 1);

            row.setKeyValue("PRKM_ToSystem_Hours", blank, sheet, rowIdx, 1);
        }

        // Water Quality Sheet
        sheet = wb.getSheet("WQP Report");
        cal.setTime(reportDate);
        sheet.getRow(2).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(3).getCell(2).setCellValue(year);

        Calendar cal3 = cal.getInstance();
        cal3.setTime(reportDate);
        cal3.set(Calendar.DAY_OF_MONTH, 1);
        int WQSidx = 12;
        for(int i=0; i < numDaysInMonth; i++)
        {
            sheet.getRow(WQSidx).getCell(0).setCellValue(excelShortDateFormat.format(cal3.getTime()));
            sheet.getRow(WQSidx).getCell(4).setCellValue("PO");
            sheet.getRow(WQSidx).getCell(5).setCellValue("PO");
            sheet.getRow(WQSidx).getCell(6).setCellValue("PO");
            WQSidx++;
            cal3.add(Calendar.DAY_OF_MONTH, 1);
        }

        Integer curr_Day = -1;
        //Integer wqRowIdx = 12;
        Double Max_Turb = 0.0;
        for (ObjectDatasetWrapper.Row row : WQData) {

            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Integer ToSystem = (Integer) row.getKeyValue("Valve_Closed");
            Double Turb = (Double) row.getKeyValue("PRKM_Turb");
            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = 12 + (day - 1);

            //If we have moved to the next day reset Max_turb and change our current day to our new day to check
            if (!day.equals(curr_Day)) {
                Max_Turb = 0.0;
                curr_Day = day;
                sheet.getRow(rowIdx).getCell(0).setCellValue(excelShortDateFormat.format(t_stamp));
//                sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
//                sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
//                sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
            }
            if (ToSystem.equals(1) && Turb > Max_Turb) {
                Max_Turb = Turb;
                sheet.getRow(rowIdx).getCell(5).setCellValue(Max_Turb);
                row.setKeyValue("PRKM_pH", blank, sheet, rowIdx, 6);
                row.setKeyValue("PRKM_Water_Temp", blank, sheet, rowIdx, 4);

            }

        }

        // Disinfection Report
        sheet = wb.getSheet("DISINFECTION REPORT");
        cal.setTime(reportDate);
        sheet.getRow(2).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(3).getCell(2).setCellValue(year);

        Calendar cal2 = cal.getInstance();
        cal2.setTime(reportDate);
        cal2.set(Calendar.DAY_OF_MONTH, 1);
        int tsIDX = 13;
        for(int i=0; i < numDaysInMonth; i++)
        {
            if (tsIDX < 28) {
                sheet.getRow(tsIDX).getCell(1).setCellValue(excelShortDateFormat.format(cal2.getTime()));
                sheet.getRow(tsIDX).getCell(2).setCellValue("PO");
            }
            if (tsIDX >= 28) {
                Integer IdxTmp = tsIDX - 15;
                sheet.getRow(IdxTmp).getCell(5).setCellValue(excelShortDateFormat.format(cal2.getTime()));
                sheet.getRow(IdxTmp).getCell(6).setCellValue("PO");
            }
            cal2.add(Calendar.DAY_OF_MONTH, 1);
            tsIDX++;
        }

        Double Min_Cl2 = 10.0;
        Integer MinCl2Idx = 13;
        Integer day2 = -1;
        //Iterate through all rows to determine the lowest CL2 amount while the well was running
        //Skips the first two rows after the valve goes from 0 to 1 resets the row skip every time a 0 is read
        for (ObjectDatasetWrapper.Row row : WQData) {

            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double CL2 = (Double) row.getKeyValue("PRKM_CL2_Res", 0.0);
            Double PRKM_Flow = (Double) row.getKeyValue("PRKM_Flow", 0.0);
            Double DIVD_Flow = (Double) row.getKeyValue("DIVD_Flow", 0.0);
            Integer ToSystem = (Integer) row.getKeyValue("Valve_Closed", 0);

            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            MinCl2Idx = 13 + (day - 1);

            if (!day.equals(day2)) {
                Min_Cl2 = 10.0;
                day2 = day;
                if (MinCl2Idx < 28) {
                    sheet.getRow(MinCl2Idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
//                    sheet.getRow(MinCl2Idx).getCell(2).setCellValue("PO");
                }
                if (MinCl2Idx >= 28) {
                    Integer IdxTmp = MinCl2Idx - 15;
                    sheet.getRow(IdxTmp).getCell(5).setCellValue(excelShortDateFormat.format(t_stamp));
//                    sheet.getRow(IdxTmp).getCell(6).setCellValue("PO");
                }
            }
            if (day.equals(day2)) {
                if (ToSystem == 1 && Min_Cl2 > CL2) {
                    Min_Cl2 = CL2;
                    if (MinCl2Idx < 28) {
                        sheet.getRow(MinCl2Idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                        sheet.getRow(MinCl2Idx).getCell(2).setCellValue(Min_Cl2);
                    }
                    if (MinCl2Idx >= 28) {
                        Integer IdxTmp = MinCl2Idx - 15;
                        sheet.getRow(IdxTmp).getCell(5).setCellValue(excelShortDateFormat.format(t_stamp));
                        sheet.getRow(IdxTmp).getCell(6).setCellValue(Min_Cl2);
                    }
                }

            }
        }

        //Sequence 1 Sheet
        sheet = wb.getSheet("SEQUENCE 1");
        cal.setTime(reportDate);
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);

        Calendar cal4 = cal.getInstance();
        cal4.setTime(reportDate);
        cal4.set(Calendar.DAY_OF_MONTH, 1);
        int DSSidx = 14;
        for(int i=0; i < numDaysInMonth; i++)
        {

            sheet.getRow(DSSidx).getCell(2).setCellValue("PO");
            sheet.getRow(DSSidx).getCell(4).setCellValue("PO");
            sheet.getRow(DSSidx).getCell(5).setCellValue("PO");
            sheet.getRow(DSSidx).getCell(6).setCellValue("PO");
            DSSidx++;
            cal4.add(Calendar.DAY_OF_MONTH, 1);
        }

        //Max flow values here
        Double maxFlow = 0.0;
        Integer currDay = -1;
        //Value to get to the correct start point to input data
        Integer offSet = 14;
        final int MAXFLOWCOL = 2, CL2COLL = 4, WATERPHCOLL = 5, WATERTEMPCOLL = 6;
        for (ObjectDatasetWrapper.Row row : WQData) {

            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double CL2 = (Double) row.getKeyValue("PRKM_CL2_Res", 0.0);
            Double PRKM_Flow = (Double) row.getKeyValue("PRKM_Flow", 0.0);
            Double DIVD_Flow = (Double) row.getKeyValue("DIVD_Flow", 0.0);
            Double waterTemp = (Double) row.getKeyValue("PRKM_Water_Temp", 0.0);
            Double waterPh = (Double) row.getKeyValue("PRKM_pH", 0.0);
            Integer ToSystem = (Integer) row.getKeyValue("Valve_Closed", 0);

            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = offSet + (day - 1);

            if (!day.equals(currDay)) {
                maxFlow = 0.0;
                currDay = day;
//                sheet.getRow(rowIdx).getCell(MAXFLOWCOL).setCellValue("PO");
//                sheet.getRow(rowIdx).getCell(CL2COLL).setCellValue("PO");
//                sheet.getRow(rowIdx).getCell(WATERPHCOLL).setCellValue("PO");
//                sheet.getRow(rowIdx).getCell(WATERTEMPCOLL).setCellValue("PO");
            }
            if ((PRKM_Flow + DIVD_Flow) > maxFlow && ToSystem == 1) {
                maxFlow = PRKM_Flow + DIVD_Flow;
                sheet.getRow(rowIdx).getCell(MAXFLOWCOL).setCellValue(maxFlow);
                sheet.getRow(rowIdx).getCell(CL2COLL).setCellValue(CL2);
                sheet.getRow(rowIdx).getCell(WATERPHCOLL).setCellValue(waterPh);
                sheet.getRow(rowIdx).getCell(WATERTEMPCOLL).setCellValue(waterTemp);
            }
        }
        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getQuinnsSheet(Date reportDate, ObjectDatasetWrapper FiveMinData, ObjectDatasetWrapper rackResults, ObjectDatasetWrapper WQData) throws Exception {

        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);

        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("Quinns Monthly Report.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);

        //EPA Monthly Summary
        Map EPA_params = new HashMap();
        EPA_params.put("sheet_name", "EPA Monthly Summary");
        EPA_params.put("cl2_column", "AI0533_EffluentCL2");
        EPA_params.put("date", "t_stamp");
        //EPA Monthly Summary
        EPA_Summary(FiveMinData, wb, month, year, EPA_params);

        // Turbidity Worksheet
        Map turb_params = new HashMap();
        turb_params.put("sheet_name", "Turbidity Daily Data Sheet");
        turb_params.put("date", "t_stamp");
        turb_params.put("12AM_turb", "QJ_12AM_Turb");
        turb_params.put("4AM_turb", "QJ_4AM_Turb");
        turb_params.put("8AM_turb", "QJ_8AM_Turb");
        turb_params.put("12PM_turb", "QJ_12PM_Turb");
        turb_params.put("4PM_turb", "QJ_4PM_Turb");
        turb_params.put("8PM_turb", "QJ_8PM_Turb");
        Turbidity_Sheet(WQData, wb, month, year, turb_params);

        //DI Testing--------------------------------------------------------------------------------------------------------------------------------------------
        //Unit(1) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "1", month, year);
        //Unit(2) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "2", month, year);
        //Unit(3) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "3", month, year);
        //Unit(4) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "4", month, year);

        //Operational Worksheet---------------------------------------------------------------------------------------------------------------------------------
        Map op_params = new HashMap();
        op_params.put("cl2_column", "AI0533_EffluentCL2");
        op_params.put("flow1", "FI0510_Fairway");
        op_params.put("flow2", "FI0520_BootHill");
        op_params.put("water_temp", "TI3008_RawWaterTemp");
        op_params.put("clearwell_level", "LI0510_ClearWellLevel");
        op_params.put("plant_running", "PlantRunning");
        op_params.put("date", "t_stamp");
        Quinn_Operational_Sheet(FiveMinData, wb, op_params);

        // Disinfection Report----------------------------------------------------------------------------------------------------------------------------------
        Quinn_Disinfection(FiveMinData, wb, month, year);

        //Sequence Dates ---------------------------------------------------------------------------------------------------------------------------------------
        Quinn_Sequence(cal, year, wb);

        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getCreekside(Date reportDate, ObjectDatasetWrapper fiveMinData, ObjectDatasetWrapper turbData, ObjectDatasetWrapper hours) throws Exception {

        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);

        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("CreeksideWTP.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);

        //EPA Monthly Summary
        Map EPA_params = new HashMap();
        EPA_params.put("sheet_name", "EPA Monthly Summary");
        EPA_params.put("cl2_column", "PRKM_CL2_Res");
        EPA_params.put("date", "t_stamp");
        //EPA Monthly Summary
        EPA_SummaryCreekside(fiveMinData, wb, month, year, EPA_params, reportDate);

        Map turb_params = new HashMap();
        turb_params.put("sheet_name", "Turbidity Data");
        turb_params.put("date", "t_stamp");
        turb_params.put("12AM_turb", "PRKM_12AM_Turb");
        turb_params.put("4AM_turb", "PRKM_4AM_Turb");
        turb_params.put("8AM_turb", "PRKM_8AM_Turb");
        turb_params.put("12PM_turb", "PRKM_12PM_Turb");
        turb_params.put("4PM_turb", "PRKM_4PM_Turb");
        turb_params.put("8PM_turb", "PRKM_8PM_Turb");
        Turbidity_SheetCreekside(turbData, wb, month, year, turb_params, hours, reportDate);


        //Operational Worksheet---------------------------------------------------------------------------------------------------------------------------------
        Map op_params = new HashMap();
        op_params.put("cl2_column", "PRKM_CL2_res");
        op_params.put("flow1", "PRKM_Flow");
        op_params.put("flow2", "DIVD_Flow");
        op_params.put("water_temp", "PRKM_Water_Temp");
        op_params.put("PRKM_pH", "PRKM_pH");
        op_params.put("date", "t_stamp");
        Creekside_Operational_Sheet(fiveMinData, wb, op_params);

        // Disinfection Report----------------------------------------------------------------------------------------------------------------------------------
        Creekside_Disinfection(fiveMinData, wb);
//
//        //Sequence Dates ---------------------------------------------------------------------------------------------------------------------------------------
//        Quinn_Sequence(cal, year, wb);

        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getJSSD(Date reportDate, ObjectDatasetWrapper fiveMinData, ObjectDatasetWrapper turbData, ObjectDatasetWrapper hours) throws Exception {

        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);

        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("JSSD_Monthly.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);

        XSSFSheet sheet = wb.getSheet("Turb Data");
//        sheet.getRow(2).getCell(2).setCellValue(reportDate);
//        sheet.getRow(3).getCell(2).setCellValue(year_in);
        int prevDay=0;

        for (ObjectDatasetWrapper.Row row : turbData) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double turb12AM = (Double) row.getKeyValue("jssd_12AM_Turb");
            Double turb4AM = (Double) row.getKeyValue("jssd_4AM_turb");
            Double turb8AM = (Double) row.getKeyValue("jssd_8AM_turb");
            Double turb12PM = (Double) row.getKeyValue("jssd_12PM_turb");
            Double turb4PM = (Double) row.getKeyValue("jssd_4PM_turb");
            Double turb8PM = (Double) row.getKeyValue("jssd_8PM_turb");
            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            if (day - prevDay > 1) {
                cal.add(Calendar.DAY_OF_MONTH, -1);
                Date t_stamp2 = cal.getTime();

                Integer rowIdx = 11 + (day - 2);
                sheet.getRow(rowIdx).getCell(0).setCellValue(excelShortDateFormat.format(t_stamp2));
                sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(3).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(7).setCellValue("PO");
            }

            cal.set(Calendar.DAY_OF_MONTH, day + 1);
            Integer rowIdx = 11 + (day - 1);

           sheet.getRow(rowIdx).getCell(0).setCellValue(excelShortDateFormat.format(t_stamp));
            if (turb4AM == null) {
                sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(2).setCellValue(turb4AM);
            }
            if (turb8AM == null) {
                sheet.getRow(rowIdx).getCell(3).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(3).setCellValue(turb8AM);
            }
            if (turb12PM == null) {
                sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(4).setCellValue(turb12PM);
            }
            if (turb4PM == null) {
                sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(5).setCellValue(turb4PM);
            }
            if (turb8PM == null) {
                sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(6).setCellValue(turb8PM);
            }
            if (turb12AM == null) {
                sheet.getRow(rowIdx).getCell(7).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(7).setCellValue(turb12AM);
            }


            prevDay = day;
        }

        sheet = wb.getSheet("Disinfection Report");

        Double Min_Cl2 = 10.0;
        final Integer MINCL2OFFSET = 13;
        Integer MinCl2Idx = 0;
        Integer day2 = -1;

        //Iterate through all rows to determine the lowest CL2 amount while the well was running
        for (ObjectDatasetWrapper.Row row : fiveMinData) {

            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double CL2 = (Double) row.getKeyValue("jssd_eff_cl2", 0.0);
            Double flow = (Double) row.getKeyValue("jssd_eff_flow", 0.0);
            if(flow != null) {
                cal.setTime(t_stamp);
                Integer day = cal.get(Calendar.DAY_OF_MONTH);
                MinCl2Idx = MINCL2OFFSET + (day - 1);

                if (!day.equals(day2)) {
                    Min_Cl2 = 10.0;
                    day2 = day;
                }
                if (day.equals(day2)) {

                    if (flow > 20 && Min_Cl2 > CL2) {
                        Min_Cl2 = CL2;
                        if (MinCl2Idx < 28) {
                            sheet.getRow(MinCl2Idx).getCell(2).setCellValue(Min_Cl2);
                        }
                        if (MinCl2Idx >= 28) {
                            Integer IdxTmp = MinCl2Idx - 15;
                            sheet.getRow(IdxTmp).getCell(6).setCellValue(Min_Cl2);
                        }
                    }
                    if(MinCl2Idx < 28) {
                        sheet.getRow(MinCl2Idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                    }
                    else if(MinCl2Idx >= 28)
                    {
                        Integer IdxTmp = MinCl2Idx - 15;
                        sheet.getRow(IdxTmp).getCell(5).setCellValue(excelShortDateFormat.format(t_stamp));
                    }

                }
            }
        }

        sheet = wb.getSheet("Operational Worksheet");
        int rowIdx = 8;
        int rowOffset = 0;
        Integer curHour = -1;
        Double maxFlow = 0.00;

        for (ObjectDatasetWrapper.Row row : fiveMinData) {
            //Load Values from the tables
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double cl2Res = (Double) row.getKeyValue("jssd_eff_cl2");
            Double flow = (Double) row.getKeyValue("jssd_eff_flow");
            Double waterTemp = (Double) row.getKeyValue("jssd_inf_water_temp");
            Double level = (Double) row.getKeyValue("jssd_fw_level");
            Double waterTempF = waterTemp * 1.8 + 32;
            Double pmPH = (Double) row.getKeyValue("jssd_ph");

            if(flow != null) {
                //Offsets for storing in the correct area
                cal.setTime(t_stamp);
                Integer day = cal.get(Calendar.DAY_OF_MONTH);
                Integer hour = cal.get(Calendar.HOUR_OF_DAY);

                //If its a new day populate each row for that day on an hourly basis
                if (!curHour.equals(hour)) {
                    //reset values and drop the index to the next line
                    rowOffset = rowIdx + ((day-1) * 24) + hour;
                    maxFlow = 0.00;
                    curHour = hour;
                    //Create a Calendar object to set the new hour
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(t_stamp);
                    calendar.set(Calendar.MINUTE, 0);
                    calendar.set(Calendar.SECOND, 0);
                    sheet.getRow(rowOffset).getCell(1).setCellValue(operationalDateFormat.format(calendar.getTime()));
                }
                //If PM Well Flow > 100 and the flow is larger than the current High Flow for the day store it
                if (flow > 100 && maxFlow < flow) {
                    sheet.getRow(rowOffset).getCell(3).setCellValue(flow);
                    sheet.getRow(rowOffset).getCell(4).setCellValue(waterTemp);
                    sheet.getRow(rowOffset).getCell(5).setCellValue(level);
                    sheet.getRow(rowOffset).getCell(6).setCellValue(cl2Res);
                }
            }

        }
//
//        //Sequence Dates ---------------------------------------------------------------------------------------------------------------------------------------
//        Quinn_Sequence(cal, year, wb);

        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getCreeksideUVDaily(Date reportDate, ObjectDatasetWrapper runHours, ObjectDatasetWrapper totalProd, ObjectDatasetWrapper redData, ObjectDatasetWrapper offSpecData) throws Exception {

        Calendar cal = Calendar.getInstance();

        Calendar cal2 = Calendar.getInstance();
        cal2.setTime(reportDate);
        cal2.set(Calendar.DAY_OF_MONTH, cal2.getActualMaximum(Calendar.DAY_OF_MONTH));

        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("CKSD_DailyUV.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);
        cal.set(Calendar.DAY_OF_MONTH, 1);

        XSSFSheet sheet = wb.getSheet("Summary");
        sheet.getRow(1).getCell(5).setCellValue(displayFormat.format(cal.getTime()) + " to " + displayFormat.format(cal2.getTime()));

        int rowIdx = 9;
        for (ObjectDatasetWrapper.Row row : runHours) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Calendar c = Calendar.getInstance();
            c.setTime(t_stamp);
            int day = c.get(Calendar.DAY_OF_MONTH);
            Double runTime = (Double) row.getKeyValue("PRKM_ToSystem_Hours");
            sheet.getRow(rowIdx+day).getCell(2).setCellValue(runTime);
        }
        rowIdx = 9;
        for (ObjectDatasetWrapper.Row row : totalProd) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Calendar c = Calendar.getInstance();
            c.setTime(t_stamp);
            int day = c.get(Calendar.DAY_OF_MONTH);
            Double total = (Double) row.getKeyValue("WEL_PRKM_PM_Daily");
            sheet.getRow(rowIdx+day).getCell(4).setCellValue(total / 1000);
        }
        rowIdx = 9;
        for (ObjectDatasetWrapper.Row row : redData) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Calendar c = Calendar.getInstance();
            c.setTime(t_stamp);
            int day =  c.get(Calendar.DAY_OF_MONTH);
            Double minRed = (Double) row.getKeyValue("minRed");
            Date minTime = (Date) row.getKeyValue("minTime");
            Double minFlow = (Double) row.getKeyValue("minFlow");
            Double minUVT = (Double) row.getKeyValue("minUVT");

            if(minRed == 999.0) {
                sheet.getRow(rowIdx + day).getCell(6).setCellValue("PO");
                sheet.getRow(rowIdx + day).getCell(12).setCellValue("PO");
                sheet.getRow(rowIdx + day).getCell(15).setCellValue("PO");
                sheet.getRow(rowIdx + day).getCell(16).setCellValue("PO");
            }
            else {
                sheet.getRow(rowIdx + day).getCell(6).setCellValue(minTime);
                sheet.getRow(rowIdx + day).getCell(12).setCellValue(minRed);
                sheet.getRow(rowIdx + day).getCell(15).setCellValue(minFlow * .00144);
                sheet.getRow(rowIdx + day).getCell(16).setCellValue(minUVT);
            }

        }
        rowIdx = 9;
        for (ObjectDatasetWrapper.Row row: offSpecData){
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Calendar c = Calendar.getInstance();
            c.setTime(t_stamp);
            int day =  c.get(Calendar.DAY_OF_MONTH);
            Double offSpecFlow = (Double) row.getKeyValue("OffSpecFlow");
            sheet.getRow(rowIdx+day).getCell(19).setCellValue(offSpecFlow *1000 / 1000000);

        }

        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getCreeksideUVMonthly(Date reportDate, ObjectDatasetWrapper runHours, ObjectDatasetWrapper totalProd, ObjectDatasetWrapper offSpecData) throws Exception {

        Calendar cal = Calendar.getInstance();

        Calendar cal2 = Calendar.getInstance();
        cal2.setTime(reportDate);
        cal2.set(Calendar.DAY_OF_MONTH, cal2.getActualMaximum(Calendar.DAY_OF_MONTH));

        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("CKSD_UVMonthly.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);
        cal.set(Calendar.DAY_OF_MONTH, 1);

        XSSFSheet sheet = wb.getSheet("Monthly Summary Form");
        sheet.getRow(1).getCell(3).setCellValue(displayFormat.format(cal.getTime()) + " to " + displayFormat.format(cal2.getTime()));

        double totalHours = 0;
        for (ObjectDatasetWrapper.Row row : runHours) {
            Double runTime = (Double) row.getKeyValue("PRKM_ToSystem_Hours");
            totalHours += runTime;
        }
        sheet.getRow(9).getCell(3).setCellValue(totalHours);

        double totalFlow = 0;
        for (ObjectDatasetWrapper.Row row : totalProd) {
            Double total = (Double) row.getKeyValue("WEL_PRKM_PM_Daily");
            totalFlow += total;
        }
        sheet.getRow(9).getCell(5).setCellValue(totalFlow *1000/1000000);

        int offSpecEvents = 0;
        totalFlow = 0;
        for (ObjectDatasetWrapper.Row row: offSpecData){
            Double events = (Double) row.getKeyValue("OffSpecEvents");
            Double flow = (Double) row.getKeyValue("OffSpecFlow");
            offSpecEvents += events;
            totalFlow += flow;
        }
        sheet.getRow(9).getCell(8).setCellValue(offSpecEvents);
        sheet.getRow(9).getCell(11).setCellValue(totalFlow *1000/1000000);
        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getCreeksideUVOffspec(Date reportDate, ObjectDatasetWrapper offSpecData) throws Exception {

        Calendar cal = Calendar.getInstance();

        Calendar cal2 = Calendar.getInstance();
        cal2.setTime(reportDate);
        cal2.set(Calendar.DAY_OF_MONTH, cal2.getActualMaximum(Calendar.DAY_OF_MONTH));


        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("UV Off-SpecWaterLog.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);
        cal.set(Calendar.DAY_OF_MONTH, 1);

        XSSFSheet sheet = wb.getSheet("OffSpecLog");
        sheet.getRow(6).getCell(2).setCellValue(displayFormat.format(cal.getTime()));
        cal.set(Calendar.DAY_OF_MONTH, 1);
        sheet.getRow(1).getCell(2).setCellValue(displayFormat.format(cal.getTime()) + " to " + displayFormat.format(cal2.getTime()));

        int rowIdx = 10;
        for (ObjectDatasetWrapper.Row row: offSpecData) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Calendar c = Calendar.getInstance();
            c.setTime(t_stamp);
            Integer timeMins = (Integer) row.getKeyValue("timeMins");
            Double offSpecFlow = (Double) row.getKeyValue("volume");
            sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
            sheet.getRow(rowIdx).getCell(2).setCellValue(hoursMinsSecs.format(t_stamp));
            sheet.getRow(rowIdx).getCell(9).setCellValue(timeMins);
            sheet.getRow(rowIdx).getCell(11).setCellValue(offSpecFlow);
            rowIdx++;
        }
        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getQuinnsSheetMnO2(Date reportDate, ObjectDatasetWrapper FiveMinData, ObjectDatasetWrapper rackResults, ObjectDatasetWrapper WQData) throws Exception {

        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);

        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("Quinns Monthly Report MN02.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) + 1;
        Integer year = cal.get(Calendar.YEAR);

        //EPA Monthly Summary
        Map EPA_params = new HashMap();
        EPA_params.put("sheet_name", "EPA Monthly Summary");
        EPA_params.put("cl2_column", "AI0533_EffluentCL2");
        EPA_params.put("date", "t_stamp");
        //EPA Monthly Summary
        EPA_Summary(FiveMinData, wb, month, year, EPA_params);

        // Turbidity Worksheet
        Map turb_params = new HashMap();
        turb_params.put("sheet_name", "Turbidity Daily Data Sheet");
        turb_params.put("date", "t_stamp");
        turb_params.put("12AM_turb", "QJ_12AM_Turb");
        turb_params.put("4AM_turb", "QJ_4AM_Turb");
        turb_params.put("8AM_turb", "QJ_8AM_Turb");
        turb_params.put("12PM_turb", "QJ_12PM_Turb");
        turb_params.put("4PM_turb", "QJ_4PM_Turb");
        turb_params.put("8PM_turb", "QJ_8PM_Turb");
        Turbidity_Sheet(WQData, wb, month, year, turb_params);

        //DI Testing--------------------------------------------------------------------------------------------------------------------------------------------
        //Unit(1) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "1", month, year);
        //Unit(2) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "2", month, year);
        //Unit(3) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "3", month, year);
        //Unit(4) DI Testing
        Quinn_DI_Testing(FiveMinData, rackResults, wb, "4", month, year);

        //Operational Worksheet---------------------------------------------------------------------------------------------------------------------------------
        Map op_params = new HashMap();
        op_params.put("cl2_column", "AI0533_EffluentCL2");
        op_params.put("flow1", "FI0510_Fairway");
        op_params.put("flow2", "FI0520_BootHill");
        op_params.put("water_temp", "TI3008_RawWaterTemp");
        op_params.put("clearwell_level", "LI0510_ClearWellLevel");
        op_params.put("plant_running", "PlantRunning");
        op_params.put("date", "t_stamp");
        op_params.put("num_trains_online", "MN02_TrainsOnline");
        Quinn_Operational_Sheet_MN02(FiveMinData, wb, op_params);

        // Disinfection Report----------------------------------------------------------------------------------------------------------------------------------
        Quinn_Disinfection(FiveMinData, wb, month, year);

        //Sequence Dates ---------------------------------------------------------------------------------------------------------------------------------------
        Quinn_Sequence_MNO2(cal, year, wb);

        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getOgdensSheet(Date reportDate, ObjectDatasetWrapper operationalData, ObjectDatasetWrapper rackResults, ObjectDatasetWrapper WQData, ObjectDatasetWrapper turbidity,
                                          ObjectDatasetWrapper rack2Results, ObjectDatasetWrapper rack3Results, ObjectDatasetWrapper rack4Results, ObjectDatasetWrapper rack5Results,
                                          ObjectDatasetWrapper rack6Results, ObjectDatasetWrapper rack7Results, ObjectDatasetWrapper rack8Results, ObjectDatasetWrapper rack9Results) throws Exception {
        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);
        //Get an input stream to the workbook
        InputStream is = ExcelFunctions.class.getResourceAsStream("OgdenReport.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        //Pull out the Month and Year
        cal.setTime(reportDate);
        Integer month = cal.get(Calendar.MONTH) +1;
        Integer year = cal.get(Calendar.YEAR);

        Ogden_EPA_Summary(operationalData, wb, month, year);

        //Ogden Operational Sheet
        Ogden_Operational_Sheet(operationalData, wb);

        //Turbidity Sheet
        Map turb_params = new HashMap();
        turb_params.put("sheet_name", "Turbidity Daily Data Sheet");
        turb_params.put("date", "t_stamp");
        turb_params.put("12AM_turb", "CFE_12AM");
        turb_params.put("4AM_turb", "CFE_4AM");
        turb_params.put("8AM_turb", "CFE_8AM");
        turb_params.put("12PM_turb", "CFE_12PM");
        turb_params.put("4PM_turb", "CFE_4PM");
        turb_params.put("8PM_turb", "CFE_8PM");
        Turbidity_Sheet(turbidity, wb, month, year, turb_params);

        Ogden_DI(rackResults, month, year, wb, "1");
        Ogden_DI(rack2Results, month, year, wb, "2");
        Ogden_DI(rack3Results, month, year, wb, "3");
        Ogden_DI(rack4Results, month, year, wb, "4");
        Ogden_DI(rack5Results, month, year, wb, "5");
        Ogden_DI(rack6Results, month, year, wb, "6");
        Ogden_DI(rack7Results, month, year, wb, "7");
        Ogden_DI(rack8Results, month, year, wb, "8");
        Ogden_DI(rack9Results, month, year, wb, "9");

        Ogden_Disinfection(cal, operationalData, wb, month, year);

        Ogden_Sequence_1(cal, WQData, wb, year);


        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }


    private static byte[] _getGroundWaterDisinfection(Date reportDate, ObjectDatasetWrapper groundWaterData, ObjectDatasetWrapper hypoSpeed) throws Exception {
        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);
        int year = cal.get(Calendar.YEAR);
        int month = cal.get(Calendar.MONTH) + 1;
        InputStream is = ExcelFunctions.class.getResourceAsStream("GW DBPR.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        XSSFSheet sheet = wb.getSheet("Groundwater Disinfection");
        boolean firstLoop = true;
        boolean secondLoop = true;
        double MDSCTot = 0;
        double PRKMTot = 0;
        double DIVTot = 0;
        double SpiroTot1 = 0;
        double SpiroTot2 = 0;
        double SpiroTot3 = 0;
        int test = 13;
        for (ObjectDatasetWrapper.Row row : groundWaterData) {
            double MDSCTotalFlow = (double) row.getKeyValue("MDSCTotalFlow");
            double MDSCTankLevel = (double) row.getKeyValue("MDSCTankLevel");
            double PRKMTotalFlow = (double) row.getKeyValue("PRKMTotalFlow");
            double PRKMTankLevel = (double) row.getKeyValue("PRKMTankLevel");
            double DIVTotalFlow = (double) row.getKeyValue("DIVTotalFlow");
            double DIVTankLevel = (double) row.getKeyValue("DIVTankLevel");
            double SpiroFilter1TotalFlow = (double) row.getKeyValue("SpiroFilter1TotalFlow");
            double SpiroFilter2TotalFlow = (double) row.getKeyValue("SpiroFilter2TotalFlow");
            double SpiroFilter3TotalFlow = (double) row.getKeyValue("SpiroFilter3TotalFlow");
            double SpiroTankLevel = (double) row.getKeyValue("SpiroTankLevel");

            if(firstLoop)
            {
                firstLoop = false;
                MDSCTot = MDSCTotalFlow;
                PRKMTot = PRKMTotalFlow;
                DIVTot = DIVTotalFlow;
                SpiroTot1 = SpiroFilter1TotalFlow;
                SpiroTot2 = SpiroFilter2TotalFlow;
                SpiroTot3 = SpiroFilter3TotalFlow;
            }
            else if(secondLoop)
            {
                secondLoop = false;

                sheet.getRow(test).getCell(1).setCellValue(MDSCTotalFlow);
                sheet.getRow(test).getCell(2).setCellValue(MDSCTotalFlow-MDSCTot);
                sheet.getRow(test).getCell(4).setCellValue(MDSCTankLevel);

                sheet.getRow(test).getCell(9).setCellValue(DIVTotalFlow);
                sheet.getRow(test).getCell(10).setCellValue(DIVTotalFlow-DIVTot);
                sheet.getRow(test).getCell(12).setCellValue(DIVTankLevel);

                sheet.getRow(test).getCell(25).setCellValue(PRKMTotalFlow);
                sheet.getRow(test).getCell(26).setCellValue(PRKMTotalFlow-PRKMTot);
                sheet.getRow(test).getCell(28).setCellValue(PRKMTankLevel);

                sheet.getRow(test).getCell(20).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(33).setCellValue(SpiroFilter1TotalFlow);
                sheet.getRow(test).getCell(34).setCellValue(SpiroFilter1TotalFlow-SpiroTot1);
                sheet.getRow(test).getCell(36).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(41).setCellValue(SpiroFilter2TotalFlow);
                sheet.getRow(test).getCell(42).setCellValue(SpiroFilter2TotalFlow-SpiroTot2);
                sheet.getRow(test).getCell(44).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(49).setCellValue(SpiroFilter3TotalFlow);
                sheet.getRow(test).getCell(50).setCellValue(SpiroFilter3TotalFlow-SpiroTot3);
                sheet.getRow(test).getCell(52).setCellValue(SpiroTankLevel);
                test++;

            }
            else
            {
                sheet.getRow(test).getCell(1).setCellValue(MDSCTotalFlow);
                sheet.getRow(test).getCell(4).setCellValue(MDSCTankLevel);

                sheet.getRow(test).getCell(9).setCellValue(DIVTotalFlow);
                sheet.getRow(test).getCell(12).setCellValue(DIVTankLevel);

                sheet.getRow(test).getCell(25).setCellValue(PRKMTotalFlow);
                sheet.getRow(test).getCell(28).setCellValue(PRKMTankLevel);

                sheet.getRow(test).getCell(20).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(33).setCellValue(SpiroFilter1TotalFlow);
                sheet.getRow(test).getCell(36).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(41).setCellValue(SpiroFilter2TotalFlow);
                sheet.getRow(test).getCell(44).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(49).setCellValue(SpiroFilter3TotalFlow);
                sheet.getRow(test).getCell(52).setCellValue(SpiroTankLevel);
                test++;
            }
        }
        int day = 1, prevDay = 0, rowIdx = 13, count = 1;
        double MDSCSpeed = 0, PRKMSpeed = 0, DIVSpeed = 0, SpiroSpeed = 0;
        double MDSCLog = 0, PRKMLog = 0, DIVLog = 0, SpiroLog = 0;
        boolean fLoop = true;
        for (ObjectDatasetWrapper.Row row : hypoSpeed) {
            double MDSCHypoSpeed = (double) row.getKeyValue("MDSCHypoSpeed");
            double PRKMHypoSpeed = (double) row.getKeyValue("PRKMHypoSpeed");
            double DIVHypoSpeed = (double) row.getKeyValue("DIVHypoSpeed");
            double SpiroHypoSpeed = (double) row.getKeyValue("SpiroHypoSpeed");
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            cal.setTime(t_stamp);
            day = cal.get(Calendar.DAY_OF_MONTH);
            if(fLoop)
            {
                prevDay = day;
                fLoop = false;
            }

            if(day != prevDay) {
                rowIdx = 13 + prevDay-1;
                if (MDSCSpeed == 0) {
                    sheet.getRow(rowIdx).getCell(3).setCellValue(0);
                }
                else
                {
                    MDSCSpeed = (MDSCSpeed / MDSCLog) * .60;
                    sheet.getRow(rowIdx).getCell(3).setCellValue((int)MDSCSpeed + " Hz");
                }
                if(PRKMSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(27).setCellValue(0);
                }
                else
                {
                    PRKMSpeed = (PRKMSpeed / PRKMLog) * .60;
                    sheet.getRow(rowIdx).getCell(27).setCellValue((int)PRKMSpeed + " Hz");
                }
                if(DIVSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(11).setCellValue(0);
                }
                else
                {
                    DIVSpeed = (DIVSpeed / DIVLog) * .60;
                    sheet.getRow(rowIdx).getCell(11).setCellValue((int)DIVSpeed + " Hz");
                }
                if(SpiroSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(19).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(35).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(43).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(51).setCellValue(0);
                }
                else
                {
                    SpiroSpeed = (SpiroSpeed / SpiroLog) * .60;
                    sheet.getRow(rowIdx).getCell(19).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(35).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(43).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(51).setCellValue((int)SpiroSpeed + " Hz");
                }
                prevDay = day;
                MDSCSpeed = 0;
                PRKMSpeed = 0;
                DIVSpeed = 0;
                SpiroSpeed = 0;
                MDSCLog = 0;
                PRKMLog = 0;
                DIVLog = 0;
                SpiroLog = 0;

            }
            if(MDSCHypoSpeed > 1)
            {
                MDSCSpeed = MDSCSpeed + MDSCHypoSpeed;
                MDSCLog++;
            }
            if(PRKMHypoSpeed > 1)
            {
                PRKMSpeed = PRKMSpeed + PRKMHypoSpeed;
                PRKMLog++;
            }
            if(DIVHypoSpeed > 1)
            {
                DIVSpeed = DIVSpeed + DIVHypoSpeed;
                DIVLog++;
            }
            if(SpiroHypoSpeed > 1)
            {
                SpiroSpeed = SpiroSpeed + SpiroHypoSpeed;
                SpiroLog++;
            }
            if(count == hypoSpeed.getSize())
            {
                rowIdx = 13 + day-1;
                if (MDSCSpeed == 0) {
                    sheet.getRow(rowIdx).getCell(3).setCellValue(0);
                }
                else
                {
                    MDSCSpeed = (MDSCSpeed / MDSCLog) * .60;
                    sheet.getRow(rowIdx).getCell(3).setCellValue((int)MDSCSpeed + " Hz");
                }
                if(PRKMSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(27).setCellValue(0);
                }
                else
                {
                    PRKMSpeed = (PRKMSpeed / PRKMLog) * .60;
                    sheet.getRow(rowIdx).getCell(27).setCellValue((int)PRKMSpeed + " Hz");
                }
                if(DIVSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(11).setCellValue(0);
                }
                else
                {
                    DIVSpeed = (DIVSpeed / DIVLog) * .60;
                    sheet.getRow(rowIdx).getCell(11).setCellValue((int)DIVSpeed + " Hz");
                }
                if(SpiroSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(19).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(35).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(43).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(51).setCellValue(0);
                }
                else
                {
                    SpiroSpeed = (SpiroSpeed / SpiroLog) * .60;
                    sheet.getRow(rowIdx).getCell(19).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(35).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(43).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(51).setCellValue((int)SpiroSpeed + " Hz");
                }
            }
            else
            {
                count++;
            }

        }
        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getGroundWaterDisinfectionNoPM(Date reportDate, ObjectDatasetWrapper groundWaterData, ObjectDatasetWrapper hypoSpeed) throws Exception {
        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);
        int year = cal.get(Calendar.YEAR);
        int month = cal.get(Calendar.MONTH) + 1;
        InputStream is = ExcelFunctions.class.getResourceAsStream("GW DBPR No PM.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        XSSFSheet sheet = wb.getSheet("Groundwater Disinfection");
        boolean firstLoop = true;
        boolean secondLoop = true;
        double MDSCTot = 0;
        double PRKMTot = 0;
        double DIVTot = 0;
        double SpiroTot1 = 0;
        double SpiroTot2 = 0;
        double SpiroTot3 = 0;
        int test = 13;
        for (ObjectDatasetWrapper.Row row : groundWaterData) {
            double MDSCTotalFlow = (double) row.getKeyValue("MDSCTotalFlow");
            double MDSCTankLevel = (double) row.getKeyValue("MDSCTankLevel");
//            double PRKMTankLevel = (double) row.getKeyValue("PRKMTankLevel");
            double DIVTotalFlow = (double) row.getKeyValue("DIVTotalFlow");
            double DIVTankLevel = (double) row.getKeyValue("DIVTankLevel");
            double SpiroFilter1TotalFlow = (double) row.getKeyValue("SpiroFilter1TotalFlow");
            double SpiroFilter2TotalFlow = (double) row.getKeyValue("SpiroFilter2TotalFlow");
            double SpiroFilter3TotalFlow = (double) row.getKeyValue("SpiroFilter3TotalFlow");
            double SpiroTankLevel = (double) row.getKeyValue("SpiroTankLevel");

            if(firstLoop)
            {
                firstLoop = false;
                MDSCTot = MDSCTotalFlow;
                DIVTot = DIVTotalFlow;
                SpiroTot1 = SpiroFilter1TotalFlow;
                SpiroTot2 = SpiroFilter2TotalFlow;
                SpiroTot3 = SpiroFilter3TotalFlow;
            }
            else if(secondLoop)
            {
                secondLoop = false;

                sheet.getRow(test).getCell(1).setCellValue(MDSCTotalFlow);
                sheet.getRow(test).getCell(2).setCellValue(MDSCTotalFlow-MDSCTot);
                sheet.getRow(test).getCell(4).setCellValue(MDSCTankLevel);

                sheet.getRow(test).getCell(9).setCellValue(DIVTotalFlow);
                sheet.getRow(test).getCell(10).setCellValue(DIVTotalFlow-DIVTot);
                sheet.getRow(test).getCell(12).setCellValue(DIVTankLevel);

                sheet.getRow(test).getCell(20).setCellValue(SpiroTankLevel);

                sheet.getRow(test).getCell(25).setCellValue(SpiroFilter1TotalFlow);
                sheet.getRow(test).getCell(26).setCellValue(SpiroFilter1TotalFlow-SpiroTot1);
                sheet.getRow(test).getCell(28).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(33).setCellValue(SpiroFilter2TotalFlow);
                sheet.getRow(test).getCell(34).setCellValue(SpiroFilter2TotalFlow-SpiroTot2);
                sheet.getRow(test).getCell(36).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(41).setCellValue(SpiroFilter3TotalFlow);
                sheet.getRow(test).getCell(42).setCellValue(SpiroFilter3TotalFlow-SpiroTot3);
                sheet.getRow(test).getCell(44).setCellValue(SpiroTankLevel);
                test++;

            }
            else
            {
                sheet.getRow(test).getCell(1).setCellValue(MDSCTotalFlow);
                sheet.getRow(test).getCell(4).setCellValue(MDSCTankLevel);

                sheet.getRow(test).getCell(9).setCellValue(DIVTotalFlow);
                sheet.getRow(test).getCell(12).setCellValue(DIVTankLevel);

                sheet.getRow(test).getCell(20).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(25).setCellValue(SpiroFilter1TotalFlow);
                sheet.getRow(test).getCell(28).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(33).setCellValue(SpiroFilter2TotalFlow);
                sheet.getRow(test).getCell(36).setCellValue(SpiroTankLevel);
                sheet.getRow(test).getCell(41).setCellValue(SpiroFilter3TotalFlow);
                sheet.getRow(test).getCell(44).setCellValue(SpiroTankLevel);
                test++;
            }
        }
        int day = 1, prevDay = 0, rowIdx = 13, count = 1;
        double MDSCSpeed = 0, PRKMSpeed = 0, DIVSpeed = 0, SpiroSpeed = 0;
        double MDSCLog = 0, PRKMLog = 0, DIVLog = 0, SpiroLog = 0;
        boolean fLoop = true;
        for (ObjectDatasetWrapper.Row row : hypoSpeed) {
            double MDSCHypoSpeed = (double) row.getKeyValue("MDSCHypoSpeed");
            double PRKMHypoSpeed = (double) row.getKeyValue("PRKMHypoSpeed");
            double DIVHypoSpeed = (double) row.getKeyValue("DIVHypoSpeed");
            double SpiroHypoSpeed = (double) row.getKeyValue("SpiroHypoSpeed");
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            cal.setTime(t_stamp);
            day = cal.get(Calendar.DAY_OF_MONTH);
            if(fLoop)
            {
                prevDay = day;
                fLoop = false;
            }

            if(day != prevDay) {
                rowIdx = 13 + prevDay-1;
                if (MDSCSpeed == 0) {
                    sheet.getRow(rowIdx).getCell(3).setCellValue(0);
                }
                else
                {
                    MDSCSpeed = (MDSCSpeed / MDSCLog) * .60;
                    sheet.getRow(rowIdx).getCell(3).setCellValue((int)MDSCSpeed + " Hz");
                }

                if(DIVSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(11).setCellValue(0);
                }
                else
                {
                    DIVSpeed = (DIVSpeed / DIVLog) * .60;
                    sheet.getRow(rowIdx).getCell(11).setCellValue((int)DIVSpeed + " Hz");
                }
                if(SpiroSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(19).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(35).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(43).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(51).setCellValue(0);
                }
                else
                {
                    SpiroSpeed = (SpiroSpeed / SpiroLog) * .60;
                    sheet.getRow(rowIdx).getCell(19).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(27).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(35).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(43).setCellValue((int)SpiroSpeed + " Hz");
                }
                prevDay = day;
                MDSCSpeed = 0;
                DIVSpeed = 0;
                SpiroSpeed = 0;
                MDSCLog = 0;
                DIVLog = 0;
                SpiroLog = 0;

            }
            if(MDSCHypoSpeed > 1)
            {
                MDSCSpeed = MDSCSpeed + MDSCHypoSpeed;
                MDSCLog++;
            }

            if(DIVHypoSpeed > 1)
            {
                DIVSpeed = DIVSpeed + DIVHypoSpeed;
                DIVLog++;
            }
            if(SpiroHypoSpeed > 1)
            {
                SpiroSpeed = SpiroSpeed + SpiroHypoSpeed;
                SpiroLog++;
            }
            if(count == hypoSpeed.getSize())
            {
                rowIdx = 13 + day-1;
                if (MDSCSpeed == 0) {
                    sheet.getRow(rowIdx).getCell(3).setCellValue(0);
                }
                else
                {
                    MDSCSpeed = (MDSCSpeed / MDSCLog) * .60;
                    sheet.getRow(rowIdx).getCell(3).setCellValue((int)MDSCSpeed + " Hz");
                }

                if(DIVSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(11).setCellValue(0);
                }
                else
                {
                    DIVSpeed = (DIVSpeed / DIVLog) * .60;
                    sheet.getRow(rowIdx).getCell(11).setCellValue((int)DIVSpeed + " Hz");
                }
                if(SpiroSpeed == 0)
                {
                    sheet.getRow(rowIdx).getCell(19).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(27).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(35).setCellValue(0);
                    sheet.getRow(rowIdx).getCell(43).setCellValue(0);
                }
                else
                {
                    SpiroSpeed = (SpiroSpeed / SpiroLog) * .60;
                    sheet.getRow(rowIdx).getCell(19).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(27).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(35).setCellValue((int)SpiroSpeed + " Hz");
                    sheet.getRow(rowIdx).getCell(43).setCellValue((int)SpiroSpeed + " Hz");
                }
            }
            else
            {
                count++;
            }

        }
        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    //Description: completes the SewerFlows spreadsheet by populating the cells with data passed in via SewerFlows
    //Inputs: the date the report is made, A sewerflows dataset
    //Returns: a byte array that can be converted to a XLS file
    private static byte[] _getQuinnsFlows(Date reportDate, ObjectDatasetWrapper SewerFlows) throws Exception {
        //Reference the Quinns report
        InputStream is = ExcelFunctions.class.getResourceAsStream("Quinns Sewer Flows.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        XSSFSheet sheet = wb.getSheet("Daily Sewer Flows");

        //Get the date from the reportDate
        Calendar cal = Calendar.getInstance();
        cal.setTime(reportDate);
        int year = cal.get(Calendar.YEAR);
        int month = cal.get(Calendar.MONTH) + 1;
        //Initial Starting point to store data
        final int TIMECOL = 0, PLANTFLOWCOL = 1, PLATESETTLERCOL = 2, DRAINFLOWCOL = 3;
        //set the date
        sheet.getRow(0).getCell(2).setCellValue(month);
        sheet.getRow(1).getCell(2).setCellValue(year);

        for (ObjectDatasetWrapper.Row row : SewerFlows) {
            double bootHill = (double) row.getKeyValue("BootHill");
            double fairwayHills = (double) row.getKeyValue("FairwayHills");
            double plateSettler = ((double) row.getKeyValue("Solids") / 1000);
            double drainFlow = ((double) row.getKeyValue("Waste")/ 1000);

            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            cal.setTime(t_stamp);
            //cal.add(Calendar.DAY_OF_MONTH, -1);
            int day = cal.get(Calendar.DAY_OF_MONTH);
            int rowIdxOffset = 6 + day - 1;
            double plantFlow = (bootHill + fairwayHills) / 1000;

            //Store the Values to their respective rows and columns
            sheet.getRow(rowIdxOffset).getCell(TIMECOL).setCellValue(displayFormat.format(cal.getTime()));
            sheet.getRow(rowIdxOffset).getCell(PLANTFLOWCOL).setCellValue(plantFlow);
            sheet.getRow(rowIdxOffset).getCell(PLATESETTLERCOL).setCellValue(plateSettler);
            sheet.getRow(rowIdxOffset).getCell(DRAINFLOWCOL).setCellValue(drainFlow);
        }
        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }
    private static byte[] _getQuinnsMonitoringData(Date reportDate, ObjectDatasetWrapper production, ObjectDatasetWrapper rackResults, ObjectDatasetWrapper chemicals) throws Exception {
        //Reference the Quinns report
        InputStream is = ExcelFunctions.class.getResourceAsStream("QJ Monitoring.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        XSSFSheet sheet = wb.getSheet("Production");

        int idx = 9;
        int dayIdx = 1;
        int prevDay = 0;
        for(ObjectDatasetWrapper.Row row : production)
        {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double rawTurb = (Double) row.getKeyValue("rawTurb");
            Double decantTurb = (Double) row.getKeyValue("decantTurb");
            Double decantFlow = (Double) row.getKeyValue("decantFlow");
            Double mfFeedTurb = (Double) row.getKeyValue("mfFeedTurb");
            Double feedTemp = (Double) row.getKeyValue("feedTemp");
            Double strainInPress = (Double) row.getKeyValue("strainInPress");
            Double strainOutPress = (Double) row.getKeyValue("strainOutPress");
            Double strainerDP = (Double) row.getKeyValue("strainerDp");
            Double membraneFeedPress = (Double) row.getKeyValue("memFeedPress");
            Double filtratePress = (Double) row.getKeyValue("filtratePress");

            Double feedFlow = (Double) row.getKeyValue("feedFlow");
            Double filtrateFlow = (Double) row.getKeyValue("filtrateFlow");
            Double XRFlow = (Double) row.getKeyValue("XRFlow");
            Double plantRecoveryToday = (Double) row.getKeyValue("plantRecoveryToday");
            Double combinedFiltrateTurb = (Double) row.getKeyValue("combinedFiltrateTurb");
            Double clearWellLevel = (Double) row.getKeyValue("clearWellLevel");
            Double feedVolume = (Double) row.getKeyValue("feedVolume");
            Double strainerBackwashVolume = (Double) row.getKeyValue("strainerBackwashVolume");
            Double netFiltrateVolume = (Double) row.getKeyValue("netFiltrateVolume");

            Calendar cal = Calendar.getInstance();
            cal.setTime(t_stamp);
            int curDay = cal.get(Calendar.DAY_OF_MONTH);
            if(curDay != prevDay){
                dayIdx = 1;
            }
            prevDay = curDay;

            sheet.getRow(idx).getCell(0).setCellValue(dayIdx);
            sheet.getRow(idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
            sheet.getRow(idx).getCell(2).setCellValue(hoursMinsSecs.format(t_stamp));
            sheet.getRow(idx).getCell(3).setCellValue(rawTurb);
            sheet.getRow(idx).getCell(4).setCellValue(decantTurb);
            sheet.getRow(idx).getCell(5).setCellValue(decantFlow);
            sheet.getRow(idx).getCell(6).setCellValue(mfFeedTurb);
            sheet.getRow(idx).getCell(7).setCellValue(feedTemp);
            sheet.getRow(idx).getCell(8).setCellValue(strainInPress);
            sheet.getRow(idx).getCell(9).setCellValue(strainOutPress);
            sheet.getRow(idx).getCell(10).setCellValue(strainerDP);
            sheet.getRow(idx).getCell(11).setCellValue(membraneFeedPress);
            sheet.getRow(idx).getCell(12).setCellValue(filtratePress);
            sheet.getRow(idx).getCell(13).setCellValue(feedFlow);
            sheet.getRow(idx).getCell(14).setCellValue(filtrateFlow);
            sheet.getRow(idx).getCell(15).setCellValue(XRFlow);
            sheet.getRow(idx).getCell(16).setCellValue(plantRecoveryToday);
            sheet.getRow(idx).getCell(17).setCellValue(combinedFiltrateTurb);
            sheet.getRow(idx).getCell(18).setCellValue(clearWellLevel);
            sheet.getRow(idx).getCell(19).setCellValue(feedVolume);
            sheet.getRow(idx).getCell(20).setCellValue(strainerBackwashVolume);
            sheet.getRow(idx).getCell(21).setCellValue(netFiltrateVolume);
            idx++;
            dayIdx++;
        }

        //Get Rack1's sheet reset everything
        sheet = wb.getSheet("Rack 1");
        idx = 9;
        dayIdx = 1;
        prevDay = 0;

        for(ObjectDatasetWrapper.Row row : rackResults)
        {

        }
        //Run evaluator to make all in sheet formulas update with the new data
        evaluator.evaluateAll();
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    private static byte[] _getMembraneReport(Date reportDate, ObjectDatasetWrapper production, ObjectDatasetWrapper rackResults, ObjectDatasetWrapper IT_data) throws Exception {

        InputStream is = ExcelFunctions.class.getResourceAsStream("Single Rack.xlsx");
        XSSFWorkbook wb = new XSSFWorkbook(is);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        XSSFSheet sheet = wb.getSheet("Production");

        int idx = 9;
        int dayIdx = 1;
        int prevDay = 0;
        for(ObjectDatasetWrapper.Row row : production)
        {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double rawTurb = (Double) row.getKeyValue("rawWaterTurb");
            Double decantTurb = (Double) row.getKeyValue("decantTurb");
            Double decantFlow = (Double) row.getKeyValue("decantFlow");
            Double mfFeedTurb = (Double) row.getKeyValue("MFTurb");
            Double feedTemp = (Double) row.getKeyValue("feedTemp");
            Double strainInPress = (Double) row.getKeyValue("strainerInletPress");
            Double strainOutPress = (Double) row.getKeyValue("strainerOutletPress");
            Double strainerDP = (Double) row.getKeyValue("strainerDP");
            Double membraneFeedPress = (Double) row.getKeyValue("membraneFeedPress");
            Double filtratePress = (Double) row.getKeyValue("filtratePress");
            Double feedFlow = (Double) row.getKeyValue("feedFlow");
            Double filtrateFlow = (Double) row.getKeyValue("filtFlow");
            Double XRFlow = (Double) row.getKeyValue("XRFlow");
            Double plantRecoveryToday = (Double) row.getKeyValue("plantRecovery");
            Double combinedFiltrateTurb = (Double) row.getKeyValue("combinedFiltTurb");
            Double clearWellLevel = (Double) row.getKeyValue("clearWellLevel");
            Double feedVolume = (Double) row.getKeyValue("feedVolumeGal");
            Double strainerBackwashVolume = (Double) row.getKeyValue("strainerBackwashVol");
            Double netFiltrateVolume = (Double) row.getKeyValue("netFiltrateVol");

            Calendar cal = Calendar.getInstance();
            cal.setTime(t_stamp);
            int curDay = cal.get(Calendar.DAY_OF_MONTH);
            if(curDay != prevDay){
                dayIdx = 1;
            }
            prevDay = curDay;
//
            sheet.getRow(idx).getCell(0).setCellValue(dayIdx);
//            sheet.getRow(idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
            sheet.getRow(idx).getCell(1).setCellValue(t_stamp);
//            sheet.getRow(idx).getCell(2).setCellValue(hoursMinsSecs.format(t_stamp));
            sheet.getRow(idx).getCell(2).setCellValue(t_stamp);
            sheet.getRow(idx).getCell(3).setCellValue(rawTurb);
            sheet.getRow(idx).getCell(4).setCellValue(decantTurb);
            sheet.getRow(idx).getCell(5).setCellValue(decantFlow);
            sheet.getRow(idx).getCell(6).setCellValue(mfFeedTurb);
            sheet.getRow(idx).getCell(7).setCellValue(feedTemp);
            sheet.getRow(idx).getCell(8).setCellValue(strainInPress);
            sheet.getRow(idx).getCell(9).setCellValue(strainOutPress);
            sheet.getRow(idx).getCell(10).setCellValue(strainerDP);
            sheet.getRow(idx).getCell(11).setCellValue(membraneFeedPress);
            sheet.getRow(idx).getCell(12).setCellValue(filtratePress);
            sheet.getRow(idx).getCell(13).setCellValue(feedFlow);
            sheet.getRow(idx).getCell(14).setCellValue(filtrateFlow);
            sheet.getRow(idx).getCell(15).setCellValue(XRFlow);
            sheet.getRow(idx).getCell(16).setCellValue(plantRecoveryToday);
            sheet.getRow(idx).getCell(17).setCellValue(combinedFiltrateTurb);
            sheet.getRow(idx).getCell(18).setCellValue(clearWellLevel);
            sheet.getRow(idx).getCell(19).setCellValue(feedVolume);
            sheet.getRow(idx).getCell(20).setCellValue(strainerBackwashVolume);
            sheet.getRow(idx).getCell(21).setCellValue(netFiltrateVolume);
            idx++;
            dayIdx++;
        }

        //Get Rack1's sheet reset everything
        sheet = wb.getSheet("Rack");
        idx = 9;
        dayIdx = 1;
        prevDay = 0;

        for(ObjectDatasetWrapper.Row row : rackResults)
        {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
//            Double rawTurb = (Double) row.getKeyValue("rawWaterTurb");
            Integer rackProcess = (Integer) row.getKeyValue("rackProc");
            Integer rackSequence = (Integer) row.getKeyValue("rackStep");
            Double feedPressure = (Double) row.getKeyValue("feedPress");
            Double feedTemp = (Double) row.getKeyValue("feedTemp");
            Double filtTurb = (Double) row.getKeyValue("filtTurbidity");
            Double tmp = (Double) row.getKeyValue("TMP");
            Double filtPress = (Double) row.getKeyValue("filtPress");
            Double feedFlow = (Double) row.getKeyValue("feedFlow");
            Double xrFlow = (Double) row.getKeyValue("XRFlow");
            Double filtFlow = (Double) row.getKeyValue("filtFlow");
            Double rackVolToday = (Double) row.getKeyValue("rackVolToday");
            Double rackWasteToday = (Double) row.getKeyValue("rackWasteToday");
            Double rackRecToday = (Double) row.getKeyValue("rackRecToday");
            Double rackFlux = (Double) row.getKeyValue("rackFlux");
            Double rackSpecFlux = (Double) row.getKeyValue("rackSpecFlux");
            Double rackLRV = (Double) row.getKeyValue("rackLRV");
            Double rackFiltTime = (Double) row.getKeyValue("rackFiltTime");

            Calendar cal = Calendar.getInstance();
            cal.setTime(t_stamp);
            int curDay = cal.get(Calendar.DAY_OF_MONTH);
            if(curDay != prevDay){
                dayIdx = 1;
            }
            prevDay = curDay;

            sheet.getRow(idx).getCell(0).setCellValue(dayIdx);
            sheet.getRow(idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
            sheet.getRow(idx).getCell(1).setCellValue(t_stamp);
            sheet.getRow(idx).getCell(2).setCellValue(hoursMinsSecs.format(t_stamp));
            sheet.getRow(idx).getCell(2).setCellValue(t_stamp);
            sheet.getRow(idx).getCell(3).setCellValue(rackProcess);
            sheet.getRow(idx).getCell(4).setCellValue(rackSequence);
            sheet.getRow(idx).getCell(5).setCellValue(feedPressure);
            sheet.getRow(idx).getCell(6).setCellValue(feedTemp);
            sheet.getRow(idx).getCell(7).setCellValue(filtTurb);
            sheet.getRow(idx).getCell(8).setCellValue(tmp);
            sheet.getRow(idx).getCell(9).setCellValue(filtPress);
            sheet.getRow(idx).getCell(10).setCellValue(feedFlow);
            sheet.getRow(idx).getCell(11).setCellValue(xrFlow);
            sheet.getRow(idx).getCell(12).setCellValue(filtFlow);
            sheet.getRow(idx).getCell(13).setCellValue(rackVolToday);
            sheet.getRow(idx).getCell(14).setCellValue(rackWasteToday);
            sheet.getRow(idx).getCell(15).setCellValue(rackRecToday);
            sheet.getRow(idx).getCell(16).setCellValue(rackFlux);
            sheet.getRow(idx).getCell(17).setCellValue(rackSpecFlux);
            sheet.getRow(idx).getCell(18).setCellValue(rackLRV);
            sheet.getRow(idx).getCell(19).setCellValue(rackFiltTime);

            idx++;
            dayIdx++;
        }
        sheet = wb.getSheet("Rack IT Data");
        idx = 9;
        dayIdx = 1;
        prevDay = 0;
//
        for(ObjectDatasetWrapper.Row row : IT_data) {
            String date = (String) row.getKeyValue("DATE");
            String time = (String) row.getKeyValue("TIME");
            Date timestamp = (Date) row.getKeyValue("TIMESTAMP");
            Double testTime = (Double) row.getKeyValue("TESTTIME");
            Double startPSI = (Double) row.getKeyValue("PRESSATSTART");
            Double endPSI = (Double) row.getKeyValue("PRESSATEND");
            Double decayPSI = (Double) row.getKeyValue("PRESSCHANGE");
            String passFail = (String) row.getKeyValue("PASSFAIL");
////
            Calendar cal = Calendar.getInstance();
            cal.setTime(timestamp);
            int curDay = cal.get(Calendar.DAY_OF_MONTH);
            if(curDay != prevDay){
                dayIdx = 1;
            }
            prevDay = curDay;
//
//            sheet.getRow(idx).getCell(0).setCellValue(dayIdx);
            sheet.getRow(idx).getCell(1).setCellValue(date);
            sheet.getRow(idx).getCell(2).setCellValue(time);
            sheet.getRow(idx).getCell(3).setCellValue(testTime);
            sheet.getRow(idx).getCell(4).setCellValue(testTime);
            sheet.getRow(idx).getCell(5).setCellValue(startPSI);
            sheet.getRow(idx).getCell(6).setCellValue(endPSI);
//
//
            sheet.getRow(idx).getCell(8).setCellValue(decayPSI);
            sheet.getRow(idx).getCell(9).setCellValue(passFail);
        }


        //Run evaluator to make all in sheet formulas update with the new data
 //       evaluator.evaluateAll();
        wb.setForceFormulaRecalculation(true);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        return bos.toByteArray();
    }

    //----------------------------------------------Helper Functions----------------------------------------------------

    //Description: Completes the DI Testing Sheets for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    private static void Quinn_DI_Testing(ObjectDatasetWrapper FiveMinData, ObjectDatasetWrapper rackResults, XSSFWorkbook wb_in, String rack_number, int month_in, int year_in) {
        //Set the date at the top of the sheet
        XSSFSheet sheet = wb_in.getSheet("Unit(" + rack_number + ") DI Testing");
        sheet.getRow(4).getCell(2).setCellValue(month_in);
        sheet.getRow(5).getCell(2).setCellValue(year_in);

        final int MINFLOW = 100;
        Calendar cal = Calendar.getInstance();

        for (ObjectDatasetWrapper.Row row : rackResults) {
            Date entryTS = (Date) row.getKeyValue("t_stamp");
            Calendar entryCal = Calendar.getInstance();
            entryCal.setTime(entryTS);
            int entryDay = entryCal.get(Calendar.DAY_OF_MONTH);

            Date rackTS = (Date) row.getKeyValue("Rack" + rack_number + "TS");
            String rackResult = (String) row.getKeyValue("Rack" + rack_number + "Result");
            cal.setTime(rackTS);
            //Get the hour and day that this check was ran
            Integer rackTSDAY = cal.get(Calendar.DAY_OF_MONTH);
            Integer rackTSMonth = cal.get(Calendar.MONTH) + 1;
            //Throw out entries that aren't from the specified month
            if (rackTSMonth.equals(month_in)) {
                Boolean flowOkay = false;
                Integer rowIdx = 14 + rackTSDAY - 1;

                int resultCheck = entryDay - rackTSDAY;

                //Check the FiveMinData table for a timestamp that matches the DI test
                //if there is one then check that the flow is greater than 100
                for (ObjectDatasetWrapper.Row five_Row : FiveMinData) {
                    //Create a calendar object for our fiveMinData table
                    Date t_stamp_five_min = (Date) five_Row.getKeyValue("t_stamp");
                    Calendar tempCal = Calendar.getInstance();
                    tempCal.setTime(t_stamp_five_min);
                    Integer fiveMinDay = tempCal.get(Calendar.DAY_OF_MONTH);
                    //Check to see if there is a time stamp in FiveMin that matches our Rack table in days, hours and minutes
                    //Check the day to make sure the flow never went below 100
                    if (fiveMinDay.equals(rackTSDAY)) {
                        Double fairwayFlow = (Double) five_Row.getKeyValue("FI0510_Fairway");
                        Double boothillFlow = (Double) five_Row.getKeyValue("FI0520_BootHill");
                        //Grab the two flow values and check if they are above the minimum required to validate the rackResult
                        if (fairwayFlow + boothillFlow > MINFLOW) {
                            flowOkay = true;
                            break;
                        } else {
                            flowOkay = false;
                        }
                    }
                }
                //If the flow total found in the code above is higher than our min
                //And the date between the two timestamps in rackResults isnt more than a day different, store the value;
                sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(rackTS));
                if (flowOkay && resultCheck <= 1) {
                    if (rackResult.equals("PASS")) {
                        sheet.getRow(rowIdx).getCell(2).setCellValue("Y");
                    } else if (rackResult.equals("FAIL")) {
                        sheet.getRow(rowIdx).getCell(2).setCellValue("N");
                    }
                }
                //If the flow total is to low, put a description in
                else if (!flowOkay) {
                    sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
                }
            }
        }

    }

    //Description: Completes the EPA_Summary sheet for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    //Map parameters: sheet_name, cl2_column, date
    private static void EPA_Summary(ObjectDatasetWrapper FiveMinData, XSSFWorkbook wb_in, int month_in, int year_in, Map parameters) {
        XSSFSheet sheet = wb_in.getSheet((String) parameters.get("sheet_name"));
        sheet.getRow(4).getCell(5).setCellValue(month_in);
        sheet.getRow(5).getCell(5).setCellValue(year_in);
        Calendar cal = Calendar.getInstance();

        int offsetIndex = 29;
        final int COL2RESET = 10, COL3RESET = 20, COL1 = 5, COL2 = 24, COL3 = 43;
        Integer curDay = -1;
        double minRes = 10.00;

        //POE min residual disinfectant residual criteria
        for (ObjectDatasetWrapper.Row row : FiveMinData) {
            Double CL2 = (Double) row.getKeyValue((String) parameters.get("cl2_column"));
            Date t_stamp = (Date) row.getKeyValue((String) parameters.get("date"));
            Double fairwayFlow = (Double) row.getKeyValue("FI0510_Fairway", 0.0);
            Double boothillFlow = (Double) row.getKeyValue("FI0520_Boothill", 0.0);
            Double totalFlow = fairwayFlow + boothillFlow;

            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = offsetIndex + (day - 1);

            if (!day.equals(curDay)) {
                minRes = 10.00;
                curDay = day;
            }
            if (day.equals(curDay)) {
                //If there is a lower residual store it for that day
                if(CL2 != null) {
                    if (minRes > CL2 && totalFlow > 20) {
                        minRes = CL2;
                        //check if the value goes in the center column
                        if (rowIdx > 38 && rowIdx < 49) {
                            rowIdx = rowIdx - COL2RESET;
                            sheet.getRow(rowIdx).getCell(COL2).setCellValue(minRes);
                            sheet.getRow(rowIdx).getCell(20).setCellValue(excelShortDateFormat.format(t_stamp));
                        }
                        //check if it goes in the third column
                        else if (rowIdx > 48) {
                            rowIdx = rowIdx - COL3RESET;
                            sheet.getRow(rowIdx).getCell(COL3).setCellValue(minRes);
                            sheet.getRow(rowIdx).getCell(39).setCellValue(excelShortDateFormat.format(t_stamp));
                        }
                        //it must go in the first column then
                        else {
                            sheet.getRow(rowIdx).getCell(COL1).setCellValue(minRes);
                            sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                        }
                    }
                }
            }
        }
    }

    //Description: Completes the EPA_Summary sheet for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    //Map parameters: sheet_name, cl2_column, date
    private static void EPA_SummaryCreekside(ObjectDatasetWrapper FiveMinData, XSSFWorkbook wb_in, int month_in, int year_in, Map parameters, Date reportDate) {
        XSSFSheet sheet = wb_in.getSheet((String) parameters.get("sheet_name"));
        sheet.getRow(4).getCell(5).setCellValue(month_in);
        sheet.getRow(5).getCell(5).setCellValue(year_in);
        Calendar cal = Calendar.getInstance();

        int offsetIndex = 29;
        final int COL2RESET = 10, COL3RESET = 20, COL1 = 5, COL2 = 24, COL3 = 43;
        Integer curDay = -1;
        double minRes = 10.00;


        Calendar cal2 = Calendar.getInstance();
        cal2.setTime(reportDate);
        int days = cal2.get(Calendar.DAY_OF_MONTH);
        for (int i=1; i <= days; i++) {
            cal2.set(Calendar.DAY_OF_MONTH, i);
            int offset = offsetIndex + (i -1);
            if (offset > 38 && offset < 49) {
                offset = offset - COL2RESET;
                sheet.getRow(offset).getCell(COL2).setCellValue("PO");
                sheet.getRow(offset).getCell(20).setCellValue(excelShortDateFormat.format(cal2.getTime()));
            }
            //check if it goes in the third column
            else if (offset > 48) {
                offset = offset - COL3RESET;
                sheet.getRow(offset).getCell(COL3).setCellValue("PO");
                sheet.getRow(offset).getCell(39).setCellValue(excelShortDateFormat.format(cal2.getTime()));
            }
            //it must go in the first column then
            else {
                sheet.getRow(offset).getCell(COL1).setCellValue("PO");
                sheet.getRow(offset).getCell(1).setCellValue(excelShortDateFormat.format(cal2.getTime()));
            }


        }


        //POE min residual disinfectant residual criteria
        for (ObjectDatasetWrapper.Row row : FiveMinData) {
            Double CL2 = (Double) row.getKeyValue((String) parameters.get("cl2_column"));
            Date t_stamp = (Date) row.getKeyValue((String) parameters.get("date"));
            Double divide = (Double) row.getKeyValue("DIVD_Flow", 0.0);
            Double parkmeadow = (Double) row.getKeyValue("PRKM_Flow", 0.0);
            int flushValveClosed = (int) row.getKeyValue("Valve_Closed");

            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = offsetIndex + (day - 1);

            if (!day.equals(curDay)) {
                minRes = 10.00;
                curDay = day;
            }
            if (day.equals(curDay)) {
                //If there is a lower residual store it for that day
                if(CL2 != null) {
                    if (minRes > CL2 && parkmeadow > 20) {
                        minRes = CL2;
                        //check if the value goes in the center column
                        if (rowIdx > 38 && rowIdx < 49) {
                            rowIdx = rowIdx - COL2RESET;
                            sheet.getRow(rowIdx).getCell(COL2).setCellValue(minRes);
                            sheet.getRow(rowIdx).getCell(20).setCellValue(excelShortDateFormat.format(t_stamp));
                        }
                        //check if it goes in the third column
                        else if (rowIdx > 48) {
                            rowIdx = rowIdx - COL3RESET;
                            sheet.getRow(rowIdx).getCell(COL3).setCellValue(minRes);
                            sheet.getRow(rowIdx).getCell(39).setCellValue(excelShortDateFormat.format(t_stamp));
                        }
                        //it must go in the first column then
                        else {
                            sheet.getRow(rowIdx).getCell(COL1).setCellValue(minRes);
                            sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                        }
                    }
                }
            }
        }
    }

    private static void Ogden_EPA_Summary(ObjectDatasetWrapper operationalData, XSSFWorkbook wb_in, int month_in, int year_in) {
        XSSFSheet sheet = wb_in.getSheet("EPA Monthly Summary");
        sheet.getRow(4).getCell(5).setCellValue(month_in);
        sheet.getRow(5).getCell(5).setCellValue(year_in);
        Calendar cal = Calendar.getInstance();

        int offsetIndex = 29;
        final int COL2RESET = 10, COL3RESET = 20, COL1 = 5, COL2 = 24, COL3 = 43;
        Integer curDay = -1;

        //POE min residual disinfectant residual criteria
        for (ObjectDatasetWrapper.Row row : operationalData) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = offsetIndex + (day - 1);

            if (!day.equals(curDay)) {
                curDay = day;
            }
            if (day.equals(curDay)) {
                //If there is a lower residual store it for that day
                //check if the value goes in the center column
                if (rowIdx > 38 && rowIdx < 49) {
                    rowIdx = rowIdx - COL2RESET;
                    sheet.getRow(rowIdx).getCell(20).setCellValue(excelShortDateFormat.format(t_stamp));
                }
                //check if it goes in the third column
                else if (rowIdx > 48) {
                    rowIdx = rowIdx - COL3RESET;
                    sheet.getRow(rowIdx).getCell(39).setCellValue(excelShortDateFormat.format(t_stamp));
                }
                //it must go in the first column then
                else {
                    sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                }
            }
        }

    }

    //Description: Completes the turbidity sheet for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    //Map parameters: sheet_name, date, t_stamp, 12AM_turb, 4AM_turb, 8AM_turb, 12PM_turb, 4PM_turb, 8PM_turb);
    private static void Turbidity_Sheet(ObjectDatasetWrapper WQData, XSSFWorkbook wb_in, int month_in, int year_in, Map parameters) {

        XSSFSheet sheet = wb_in.getSheet((String) parameters.get("sheet_name"));
        sheet.getRow(3).getCell(2).setCellValue(month_in);
        sheet.getRow(4).getCell(2).setCellValue(year_in);
        Double blank = sheet.getRow(12).getCell(2).getNumericCellValue();
        Calendar cal = Calendar.getInstance();
        int prevDay=0;

        for (ObjectDatasetWrapper.Row row : WQData) {
            Date t_stamp = (Date) row.getKeyValue((String) parameters.get("date"));
            Double turb12AM = (Double) row.getKeyValue((String) parameters.get("12AM_turb"));
            Double turb4AM = (Double) row.getKeyValue((String) parameters.get("4AM_turb"));
            Double turb8AM = (Double) row.getKeyValue((String) parameters.get("8AM_turb"));
            Double turb12PM = (Double) row.getKeyValue((String) parameters.get("12PM_turb"));
            Double turb4PM = (Double) row.getKeyValue((String) parameters.get("4PM_turb"));
            Double turb8PM = (Double) row.getKeyValue((String) parameters.get("8PM_turb"));
            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            if (day - prevDay > 1) {
                cal.add(Calendar.DAY_OF_MONTH, -1);
                Date t_stamp2 = cal.getTime();

                Integer rowIdx = 15 + (day - 2);
                sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp2));
                sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(3).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(7).setCellValue("PO");
            }

                cal.set(Calendar.DAY_OF_MONTH, day + 1);
                Integer rowIdx = 15 + (day - 1);

                sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                if (turb4AM == null) {
                    sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
                } else {
                    sheet.getRow(rowIdx).getCell(2).setCellValue(turb4AM);
                }
                if (turb8AM == null) {
                    sheet.getRow(rowIdx).getCell(3).setCellValue("PO");
                } else {
                    sheet.getRow(rowIdx).getCell(3).setCellValue(turb8AM);
                }
                if (turb12PM == null) {
                    sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
                } else {
                    sheet.getRow(rowIdx).getCell(4).setCellValue(turb12PM);
                }
                if (turb4PM == null) {
                    sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
                } else {
                    sheet.getRow(rowIdx).getCell(5).setCellValue(turb4PM);
                }
                if (turb8PM == null) {
                    sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
                } else {
                    sheet.getRow(rowIdx).getCell(6).setCellValue(turb8PM);
                }
                if (turb12AM == null) {
                    sheet.getRow(rowIdx).getCell(7).setCellValue("PO");
                } else {
                    sheet.getRow(rowIdx).getCell(7).setCellValue(turb12AM);
                }


            prevDay = day;
        }

    }

    //Description: Completes the turbidity sheet for Creekside
    //Inputs: See parameters
    //Returns: Nothing
    //Map parameters: sheet_name, date, t_stamp, 12AM_turb, 4AM_turb, 8AM_turb, 12PM_turb, 4PM_turb, 8PM_turb);
    private static void Turbidity_SheetCreekside(ObjectDatasetWrapper WQData, XSSFWorkbook wb_in, int month_in, int year_in, Map parameters, ObjectDatasetWrapper Hours, Date reportDate) {

        XSSFSheet sheet = wb_in.getSheet((String) parameters.get("sheet_name"));
        sheet.getRow(2).getCell(2).setCellValue(reportDate);
        sheet.getRow(3).getCell(2).setCellValue(year_in);
        Calendar cal = Calendar.getInstance();
        int prevDay=0;

        for (ObjectDatasetWrapper.Row row : WQData) {
            Date t_stamp = (Date) row.getKeyValue((String) parameters.get("date"));
            Double turb12AM = (Double) row.getKeyValue((String) parameters.get("12AM_turb"));
            Double turb4AM = (Double) row.getKeyValue((String) parameters.get("4AM_turb"));
            Double turb8AM = (Double) row.getKeyValue((String) parameters.get("8AM_turb"));
            Double turb12PM = (Double) row.getKeyValue((String) parameters.get("12PM_turb"));
            Double turb4PM = (Double) row.getKeyValue((String) parameters.get("4PM_turb"));
            Double turb8PM = (Double) row.getKeyValue((String) parameters.get("8PM_turb"));
            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            if (day - prevDay > 1) {
                cal.add(Calendar.DAY_OF_MONTH, -1);
                Date t_stamp2 = cal.getTime();

                Integer rowIdx = 14 + (day - 2);
                sheet.getRow(rowIdx).getCell(0).setCellValue(excelShortDateFormat.format(t_stamp2));
                sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(3).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
                sheet.getRow(rowIdx).getCell(7).setCellValue("PO");
            }

            cal.set(Calendar.DAY_OF_MONTH, day + 1);
            Integer rowIdx = 14 + (day - 1);

            sheet.getRow(rowIdx).getCell(0).setCellValue(excelShortDateFormat.format(t_stamp));
            if (turb4AM == null) {
                sheet.getRow(rowIdx).getCell(2).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(2).setCellValue(turb4AM);
            }
            if (turb8AM == null) {
                sheet.getRow(rowIdx).getCell(3).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(3).setCellValue(turb8AM);
            }
            if (turb12PM == null) {
                sheet.getRow(rowIdx).getCell(4).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(4).setCellValue(turb12PM);
            }
            if (turb4PM == null) {
                sheet.getRow(rowIdx).getCell(5).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(5).setCellValue(turb4PM);
            }
            if (turb8PM == null) {
                sheet.getRow(rowIdx).getCell(6).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(6).setCellValue(turb8PM);
            }
            if (turb12AM == null) {
                sheet.getRow(rowIdx).getCell(7).setCellValue("PO");
            } else {
                sheet.getRow(rowIdx).getCell(7).setCellValue(turb12AM);
            }


            prevDay = day;
        }

        for (ObjectDatasetWrapper.Row row : Hours) {
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double hour = (Double) row.getKeyValue("PRKM_ToSystem_Hours");
            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = 14 + (day - 1);
            sheet.getRow(rowIdx).getCell(1).setCellValue(hour);
        }

    }

    //Description: Completes the Operational sheet for Creekside
    //Inputs: See parameters
    //Returns: Nothing
    //Map parameters: cl2_column, flow1, flow2, water_temp, clearwell_level, plant_running, date
    private static void Creekside_Operational_Sheet(ObjectDatasetWrapper FiveMinData, XSSFWorkbook wb_in, Map op_params) {

        XSSFSheet sheet = wb_in.getSheet("Operational Worksheet");
        int rowIdx = 7;
        int rowOffset = 0;
        Integer curHour = -1;
        Double maxFlow = 0.00;
        Calendar cal = Calendar.getInstance();

        for (ObjectDatasetWrapper.Row row : FiveMinData) {
            //Load Values from the tables
            Date t_stamp = (Date) row.getKeyValue((String) op_params.get("date"));
            Double cl2Res = (Double) row.getKeyValue((String) op_params.get("cl2_column"));
            Double pmFlow = (Double) row.getKeyValue((String) op_params.get("flow1"));
            Double dvFlow = (Double) row.getKeyValue((String) op_params.get("flow2"));
            Double waterTemp = (Double) row.getKeyValue((String) op_params.get("water_temp"));
            Double waterTempF = waterTemp * 1.8 + 32;
            Double pmPH = (Double) row.getKeyValue((String) op_params.get("PRKM_pH"));

            if(pmFlow != null && dvFlow != null) {
                //Offsets for storing in the correct area
                cal.setTime(t_stamp);
                Integer day = cal.get(Calendar.DAY_OF_MONTH);
                Integer hour = cal.get(Calendar.HOUR_OF_DAY);

                //If its a new day populate each row for that day on an hourly basis
                if (!curHour.equals(hour)) {
                    //reset values and drop the index to the next line
                    rowOffset = rowIdx + ((day-1) * 24) + hour;
                    maxFlow = 0.00;
                    curHour = hour;
                    //Create a Calendar object to set the new hour
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(t_stamp);
                    calendar.set(Calendar.MINUTE, 0);
                    calendar.set(Calendar.SECOND, 0);
                    sheet.getRow(rowOffset).getCell(1).setCellValue(operationalDateFormat.format(calendar.getTime()));
                }
                //If PM Well Flow > 100 and the flow is larger than the current High Flow for the day store it
                if (pmFlow > 100 && maxFlow < pmFlow) {
                    sheet.getRow(rowOffset).getCell(3).setCellValue(pmFlow);
                    sheet.getRow(rowOffset).getCell(4).setCellValue(dvFlow);
                    sheet.getRow(rowOffset).getCell(6).setCellValue(pmPH);
                    sheet.getRow(rowOffset).getCell(7).setCellValue(waterTempF);
                    sheet.getRow(rowOffset).getCell(8).setCellValue(cl2Res);
                }
            }

        }
    }

    //Description: Completes the Operational sheet for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    //Map parameters: cl2_column, flow1, flow2, water_temp, clearwell_level, plant_running, date
    private static void Quinn_Operational_Sheet_MN02(ObjectDatasetWrapper FiveMinData, XSSFWorkbook wb_in, Map op_params) {

        XSSFSheet sheet = wb_in.getSheet("Operational Worksheet");
        int rowIdx = 7;
        Integer curHour = -1;
        Double maxFlow = 0.00;
        Calendar cal = Calendar.getInstance();

        for (ObjectDatasetWrapper.Row row : FiveMinData) {
            //Load Values from the tables
            Date t_stamp = (Date) row.getKeyValue((String) op_params.get("date"));
            Double cl2Res = (Double) row.getKeyValue((String) op_params.get("cl2_column"));
            Double fairwayFlow = (Double) row.getKeyValue((String) op_params.get("flow1"));
            Double boothillFlow = (Double) row.getKeyValue((String) op_params.get("flow2"));
            Double waterTemp = (Double) row.getKeyValue((String) op_params.get("water_temp"));
            Double clearwellLevel = (Double) row.getKeyValue((String) op_params.get("clearwell_level"));
            Integer plantRunning = (Integer) row.getKeyValue((String) op_params.get("plant_running"));
            Integer numTrainsOnline = (Integer) row.getKeyValue((String) op_params.get("num_trains_online"));

            if(fairwayFlow != null && boothillFlow != null) {
                //Offsets for storing in the correct area
                cal.setTime(t_stamp);
                Integer day = cal.get(Calendar.DAY_OF_MONTH);
                Integer hour = cal.get(Calendar.HOUR_OF_DAY);

                //If it a new day populate each row for that day on an hourly basis
                if (!curHour.equals(hour)) {
                    //reset values and drop the index to the next line
                    rowIdx++;
                    maxFlow = 0.00;
                    curHour = hour;
                    //Create a Calendar object to set the new hour
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(t_stamp);
                    calendar.set(Calendar.MINUTE, 0);
                    calendar.set(Calendar.SECOND, 0);
                    sheet.getRow(rowIdx).getCell(1).setCellValue(operationalDateFormat.format(calendar.getTime()));
                }
                //If Fairway and Boothills total flow is greater than 100 and the flow value is higher than our current for the day, store it.
                if (((fairwayFlow + boothillFlow) > 100 && maxFlow < (fairwayFlow + boothillFlow))) {
                    maxFlow = fairwayFlow + boothillFlow;
                    sheet.getRow(rowIdx).getCell(3).setCellValue(maxFlow);
                    sheet.getRow(rowIdx).getCell(4).setCellValue(waterTemp);
                    sheet.getRow(rowIdx).getCell(5).setCellValue(clearwellLevel);
                    sheet.getRow(rowIdx).getCell(6).setCellValue(cl2Res);
                    if(numTrainsOnline != null) {
                        sheet.getRow(rowIdx).getCell(7).setCellValue(numTrainsOnline);
                    }
                    else{
                        sheet.getRow(rowIdx).getCell(7).setCellValue(0);
                    }
                    sheet.getRow(rowIdx).getCell(2).setCellValue(day);
                }
            }

        }
    }

    //Description: Completes the Operational sheet for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    //Map parameters: cl2_column, flow1, flow2, water_temp, clearwell_level, plant_running, date
    private static void Quinn_Operational_Sheet(ObjectDatasetWrapper FiveMinData, XSSFWorkbook wb_in, Map op_params) {

        XSSFSheet sheet = wb_in.getSheet("Operational Worksheet");
        int rowIdx = 7;
        Integer curHour = -1;
        Double maxFlow = 0.00;
        Calendar cal = Calendar.getInstance();

        for (ObjectDatasetWrapper.Row row : FiveMinData) {
            //Load Values from the tables
            Date t_stamp = (Date) row.getKeyValue((String) op_params.get("date"));
            Double cl2Res = (Double) row.getKeyValue((String) op_params.get("cl2_column"));
            Double fairwayFlow = (Double) row.getKeyValue((String) op_params.get("flow1"));
            Double boothillFlow = (Double) row.getKeyValue((String) op_params.get("flow2"));
            Double waterTemp = (Double) row.getKeyValue((String) op_params.get("water_temp"));
            Double clearwellLevel = (Double) row.getKeyValue((String) op_params.get("clearwell_level"));
            Integer plantRunning = (Integer) row.getKeyValue((String) op_params.get("plant_running"));

            if(fairwayFlow != null && boothillFlow != null) {
                //Offsets for storing in the correct area
                cal.setTime(t_stamp);
                Integer day = cal.get(Calendar.DAY_OF_MONTH);
                Integer hour = cal.get(Calendar.HOUR_OF_DAY);

                //If its a new day populate each row for that day on an hourly basis
                if (!curHour.equals(hour)) {
                    //reset values and drop the index to the next line
                    rowIdx++;
                    maxFlow = 0.00;
                    curHour = hour;
                    //Create a Calendar object to set the new hour
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(t_stamp);
                    calendar.set(Calendar.MINUTE, 0);
                    calendar.set(Calendar.SECOND, 0);
                    sheet.getRow(rowIdx).getCell(1).setCellValue(operationalDateFormat.format(calendar.getTime()));
                }
                //If Fairway and Boothills total flow is greater than 100 and the flow value is higher than our current for the day, store it.
                if (((fairwayFlow + boothillFlow) > 100 && maxFlow < (fairwayFlow + boothillFlow))) {
                    maxFlow = fairwayFlow + boothillFlow;
                    sheet.getRow(rowIdx).getCell(3).setCellValue(maxFlow);
                    sheet.getRow(rowIdx).getCell(4).setCellValue(waterTemp);
                    sheet.getRow(rowIdx).getCell(5).setCellValue(clearwellLevel);
                    sheet.getRow(rowIdx).getCell(6).setCellValue(cl2Res);
                    sheet.getRow(rowIdx).getCell(2).setCellValue(day);
                }
            }

        }
    }



    //Description: Completes the Disinfection sheet for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    private static void Quinn_Disinfection(ObjectDatasetWrapper FiveMinData, XSSFWorkbook wb, int month, int year) {
        XSSFSheet sheet = wb.getSheet("Disinfection Report");
        sheet.getRow(2).getCell(2).setCellValue(month);
        sheet.getRow(3).getCell(2).setCellValue(year);
        Calendar cal = Calendar.getInstance();

        Double Min_Cl2 = 10.0;
        final Integer MINCL2OFFSET = 13;
        Integer MinCl2Idx = 0;
        Integer day2 = -1;

        //Iterate through all rows to determine the lowest CL2 amount while the well was running
        for (ObjectDatasetWrapper.Row row : FiveMinData) {

            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double CL2 = (Double) row.getKeyValue("AI0533_EffluentCL2", 0.0);
            Double fairwayFlow = (Double) row.getKeyValue("FI0510_Fairway", 0.0);
            Double boothillFlow = (Double) row.getKeyValue("FI0520_Boothill", 0.0);
            if(fairwayFlow != null && boothillFlow != null) {
                cal.setTime(t_stamp);
                Integer day = cal.get(Calendar.DAY_OF_MONTH);
                MinCl2Idx = MINCL2OFFSET + (day - 1);

                if (!day.equals(day2)) {
                    Min_Cl2 = 10.0;
                    day2 = day;
                }
                if (day.equals(day2)) {

                    if ((fairwayFlow + boothillFlow) > 20 && Min_Cl2 > CL2) {
                        Min_Cl2 = CL2;
                        if (MinCl2Idx < 28) {
                            sheet.getRow(MinCl2Idx).getCell(2).setCellValue(Min_Cl2);
                        }
                        if (MinCl2Idx >= 28) {
                            Integer IdxTmp = MinCl2Idx - 15;
                            sheet.getRow(IdxTmp).getCell(6).setCellValue(Min_Cl2);
                        }
                    }
                    if(MinCl2Idx < 28) {
                        sheet.getRow(MinCl2Idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                    }
                    else if(MinCl2Idx >= 28)
                    {
                        Integer IdxTmp = MinCl2Idx - 15;
                        sheet.getRow(IdxTmp).getCell(5).setCellValue(excelShortDateFormat.format(t_stamp));
                    }

                }
            }
        }
    }

    //Description: Completes the Disinfection sheet for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    private static void Creekside_Disinfection(ObjectDatasetWrapper FiveMinData, XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("Disinfection Report");
        Calendar cal = Calendar.getInstance();

        Double Min_Cl2 = 10.0;
        final Integer MINCL2OFFSET = 13;
        Integer MinCl2Idx = 0;
        Integer day2 = -1;

        //Iterate through all rows to determine the lowest CL2 amount while the well was running
        for (ObjectDatasetWrapper.Row row : FiveMinData) {

            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double CL2 = (Double) row.getKeyValue("PRKM_CL2_Res", 0.0);
            Double pmFlow = (Double) row.getKeyValue("PRKM_Flow", 0.0);
            Double dvFlow = (Double) row.getKeyValue("DIVD_Flow", 0.0);
            if(pmFlow != null && dvFlow != null) {
                cal.setTime(t_stamp);
                Integer day = cal.get(Calendar.DAY_OF_MONTH);
                MinCl2Idx = MINCL2OFFSET + (day - 1);

                if (!day.equals(day2)) {
                    Min_Cl2 = 10.0;
                    day2 = day;
                }
                if (day.equals(day2)) {

                    if ((pmFlow + dvFlow) > 20 && Min_Cl2 > CL2) {
                        Min_Cl2 = CL2;
                        if (MinCl2Idx < 28) {
                            sheet.getRow(MinCl2Idx).getCell(2).setCellValue(Min_Cl2);
                        }
                        if (MinCl2Idx >= 28) {
                            Integer IdxTmp = MinCl2Idx - 15;
                            sheet.getRow(IdxTmp).getCell(6).setCellValue(Min_Cl2);
                        }
                    }
                    if(MinCl2Idx < 28) {
                        sheet.getRow(MinCl2Idx).getCell(1).setCellValue(excelShortDateFormat.format(t_stamp));
                    }
                    else if(MinCl2Idx >= 28)
                    {
                        Integer IdxTmp = MinCl2Idx - 15;
                        sheet.getRow(IdxTmp).getCell(5).setCellValue(excelShortDateFormat.format(t_stamp));
                    }

                }
            }
        }
    }

    //Description: Sets the date on the Sequence sheets for Quinn's
    //Inputs: See parameters
    //Returns: Nothing
    private static void Quinn_Sequence(Calendar cal, int year, XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("Sequence 1");
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);

        sheet = wb.getSheet("Sequence 2");
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);
    }

    private static void Quinn_Sequence_MNO2(Calendar cal, int year, XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("Sequence 1");
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);

        sheet = wb.getSheet("Sequence 2");
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);

        sheet = wb.getSheet("Sequence 3");
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);

        sheet = wb.getSheet("Sequence 4");
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);
    }

    private static void Ogden_Operational_Sheet(ObjectDatasetWrapper operationalData, XSSFWorkbook wb) {
        XSSFSheet sheet = wb.getSheet("Operational Worksheet");
        int rowIdx = 8;
        int prevHour = -1;
        Calendar cal = Calendar.getInstance();

        for (ObjectDatasetWrapper.Row row : operationalData) {
            //Load Values from the tables
            Date t_stamp = (Date) row.getKeyValue("t_stamp");
            Double cl2Res = (Double) row.getKeyValue("Chlorine_Res");
            Double op_flow = (Double) row.getKeyValue("Op_Flow");
            Double waterTemp = (Double) row.getKeyValue("Op_Temp");

            //Offsets for storing in the correct area
            cal.setTime(t_stamp);
            Integer day = cal.get(Calendar.DAY_OF_MONTH);
            Integer hour = cal.get(Calendar.HOUR_OF_DAY);

            //Create a Calendar object to set the new hour
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(t_stamp);
            //If it logs earlier than on the hour add 1 to correct it
            if (prevHour == hour) {
                calendar.add(Calendar.HOUR, 1);
            }
            else {
                prevHour = hour;

                calendar.set(Calendar.MINUTE, 0);
                calendar.set(Calendar.SECOND, 0);
                sheet.getRow(rowIdx).getCell(1).setCellValue(operationalDateFormat.format(calendar.getTime()));
                sheet.getRow(rowIdx).getCell(3).setCellValue(op_flow);
                sheet.getRow(rowIdx).getCell(4).setCellValue(waterTemp);
                sheet.getRow(rowIdx).getCell(5).setCellValue(0.0);
                sheet.getRow(rowIdx).getCell(6).setCellValue(cl2Res);
                sheet.getRow(rowIdx).getCell(2).setCellValue(day);

                rowIdx++;
            }

        }
    }

    private static void Ogden_Sequence_1(Calendar cal, ObjectDatasetWrapper WQData, XSSFWorkbook wb, int year) {
        XSSFSheet sheet = wb.getSheet("Sequence 1");
        sheet.getRow(1).getCell(2).setCellValue(monthFormat.format(cal.getTime()));
        sheet.getRow(2).getCell(2).setCellValue(year);
        final int PH_COLUMN = 4;
        int rowIdx = 16;
        for (ObjectDatasetWrapper.Row row : WQData) {
            sheet.getRow(rowIdx).getCell(PH_COLUMN).setCellValue((Double) row.getKeyValue("Finished_Water_pH_Ave"));
            rowIdx++;
        }

    }

    private static void Ogden_DI(ObjectDatasetWrapper rackResults, int month, int year, XSSFWorkbook wb_in, String rack_number_in) {
        XSSFSheet sheet = wb_in.getSheet("Unit(" + rack_number_in + ") DI Testing");
        sheet.getRow(4).getCell(2).setCellValue(month);
        sheet.getRow(5).getCell(2).setCellValue(year);
        Calendar cal = Calendar.getInstance();
        int rack_number = Integer.parseInt(rack_number_in);

        for (ObjectDatasetWrapper.Row row : rackResults) {
            //Grab values needed and get the index
            Date time = (Date) row.getKeyValue("t_stamp");
            String pass_fail = (String) row.getKeyValue("Pass_Fail");
            cal.setTime(time);
            int day = cal.get(Calendar.DAY_OF_MONTH);
            Integer rowIdx = 14 + day - 1;
            //Set the date column
            sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(time));
            if (pass_fail!= null && pass_fail.equals("PASS")) {
                sheet.getRow(rowIdx).getCell(2).setCellValue("Y");
            } else {
                sheet.getRow(rowIdx).getCell(2).setCellValue("N");
            }
        }
    }

    private static void Ogden_Disinfection(Calendar cal_in, ObjectDatasetWrapper operationalData, XSSFWorkbook wb, int month, int year){
        XSSFSheet sheet = wb.getSheet("Disinfection Report");
        int day = 0, rowIdx;
        sheet.getRow(2).getCell(2).setCellValue(monthFormat.format(cal_in.getTime()));
        sheet.getRow(3).getCell(2).setCellValue(year);
        Calendar cal = Calendar.getInstance();
        for(ObjectDatasetWrapper.Row row : operationalData)
        {
            Date time = (Date) row.getKeyValue("t_stamp");
            cal.setTime(time);
            int cur_day = cal.get(Calendar.DAY_OF_MONTH);
            if(cur_day != day)
            {
                rowIdx = 14 + day -1;
                if(rowIdx < 28)
                {
                    sheet.getRow(rowIdx).getCell(1).setCellValue(excelShortDateFormat.format(time));
                }
                else
                {
                    rowIdx = rowIdx - 15;
                    sheet.getRow(rowIdx).getCell(5).setCellValue(excelShortDateFormat.format(time));
                }
                day = cur_day;
            }





        }
    }
}