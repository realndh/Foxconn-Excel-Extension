package com.foxconn;

import com.thingworx.common.RESTAPIConstants;
import com.thingworx.common.exceptions.InvalidRequestException;
import com.thingworx.data.util.InfoTableInstanceFactory;
import com.thingworx.entities.utils.ThingUtilities;
import com.thingworx.logging.LogUtilities;
import com.thingworx.metadata.FieldDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceDefinition;
import com.thingworx.metadata.annotations.ThingworxServiceParameter;
import com.thingworx.metadata.annotations.ThingworxServiceResult;
import com.thingworx.resources.Resource;
import com.thingworx.resources.queries.InfoTableFunctions;
import com.thingworx.things.Thing;
import com.thingworx.things.repository.FileRepositoryThing;
import com.thingworx.types.BaseTypes;
import com.thingworx.types.InfoTable;
import com.thingworx.types.collections.ValueCollection;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelResources extends Resource {

    private static final Logger _logger = LogUtilities.getInstance().getApplicationLogger(ExcelResources.class);
    private static final String THINGWORX_FILE_REPOSITORIES = "/Thingworx/FileRepositories/";
    private static final String MODULE = "Module";
    private static final String MATERIAL_CAD = "MaterialCAD";
    private static final int head_row_count = 2;

    private static List<Integer> cellsList = List.of(0, 1, 2, 3, 4, 5, 6, 13, 14, 15, 16, 26, 27);

    public ExcelResources() {
        // TODO Auto-generated constructor stub
    }

    @ThingworxServiceDefinition(name = "ExcelImport", description = "", category = "", isAllowOverride = false, aspects = {
            "isAsync:false"})
    @ThingworxServiceResult(name = "Result", description = "", baseType = "INFOTABLE", aspects = {
            "isEntityDataShape:true"})
    public InfoTable ExcelImport(
            @ThingworxServiceParameter(name = "header_row_count", description = "Number of header rows", baseType = "INTEGER", aspects = {
                    "isRequired:true", "defaultValue:1"}) Integer header_row_count,
            @ThingworxServiceParameter(name = "fileRepository", description = "File repository name", baseType = "THINGNAME", aspects = {
                    "isRequired:true", "thingTemplate:FileRepository"}) String fileRepository,
            @ThingworxServiceParameter(name = "path", description = "Path to file", baseType = "STRING", aspects = {
                    "isRequired:true"}) String path,
            @ThingworxServiceParameter(name = "dataShape", description = "Data shape", baseType = "DATASHAPENAME", aspects = {
                    "isRequired:true"}) String dataShape) throws Exception {

        long start = System.currentTimeMillis();

        Thing thing = ThingUtilities.findThing(fileRepository);

        if (thing == null) {
            throw new InvalidRequestException("File Repository [" + fileRepository + "] Does Not Exist", RESTAPIConstants.StatusCode.STATUS_NOT_FOUND);
        } else if (!(thing instanceof FileRepositoryThing)) {
            throw new InvalidRequestException("Thing [" + fileRepository + "] Is Not A File Repository", RESTAPIConstants.StatusCode.STATUS_NOT_FOUND);
        } else {

            FileRepositoryThing repo = (FileRepositoryThing) thing;

            InfoTable infoTable = InfoTableInstanceFactory.createInfoTableFromDataShape(dataShape);

            try (InputStream inputStream = repo.openFileForRead(path)) {

                ArrayList<FieldDefinition> orderedFields = infoTable.getDataShape().getFields().getOrderedFieldsByOrdinal();

                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                XSSFSheet sheet = workbook.getSheetAt(0);

                if (header_row_count == null) {
                    header_row_count = 1;
                }
                int rows = sheet.getLastRowNum();

                for (int i = header_row_count; i <= rows; i++) {
                    XSSFRow row = sheet.getRow(i);
                    if (row != null) {
                        int cells = row.getLastCellNum();
                        ValueCollection valueCollection = new ValueCollection();
                        valueCollection.SetBooleanValue(orderedFields.get(0).getName(), true);

                        ArrayList<String> errArrayList = new ArrayList<>();

                        for (int j = 0; j < cells; j++) {
                            FieldDefinition fieldDefinition = orderedFields.get(j + 2);
                            XSSFCell cell = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                            if (cell == null) {
                                if (fieldDefinition.isPrimaryKey()) {
                                    if (orderedFields.size() - cells <= 2) {
                                        errArrayList.add(fieldDefinition.getName() + " is required.");
                                        valueCollection.SetBooleanValue(orderedFields.get(0).getName(), false);
                                    }
                                }
                            } else {
                                try {
                                    switch (cell.getCellType()) {
                                        case STRING:
                                            valueCollection.put(fieldDefinition.getName(), BaseTypes.ConvertToPrimitive(cell.getRichStringCellValue().getString(), fieldDefinition.getBaseType()));
                                            break;
                                        case NUMERIC:
                                            valueCollection.put(fieldDefinition.getName(), BaseTypes.ConvertToPrimitive(cell.getNumericCellValue(), fieldDefinition.getBaseType()));
                                            break;
                                        case BOOLEAN:
                                            valueCollection.put(fieldDefinition.getName(), BaseTypes.ConvertToPrimitive(cell.getBooleanCellValue(), fieldDefinition.getBaseType()));
                                            break;
                                        case FORMULA:
                                            valueCollection.put(fieldDefinition.getName(), BaseTypes.ConvertToPrimitive(cell.getCellFormula(), fieldDefinition.getBaseType()));
                                            break;
                                    }
                                } catch (Exception e) {
                                    if (orderedFields.size() - cells <= 2) {
                                        errArrayList.add("Field : " + fieldDefinition.getName() + ", Value : " + cell.getStringCellValue() + ", Error : " + e.getMessage());
                                        valueCollection.SetBooleanValue(orderedFields.get(0).getName(), false);
                                    }
                                }

                            }
                        }
                        valueCollection.SetStringValue(orderedFields.get(1).getName(), StringUtils.join(errArrayList, System.getProperty("line.separator")));
                        infoTable.addRow(valueCollection);
                    }
                }

            } catch (Exception ex) {
                throw new InvalidRequestException(ex.getMessage(), RESTAPIConstants.StatusCode.STATUS_NOT_FOUND);
            }

            long end = System.currentTimeMillis();
            int time = (int) ((end - start) / 1000);

            _logger.warn("{} : {}", "ExcelImport Execution time : ", time);

            return infoTable;
        }
    }

    @ThingworxServiceDefinition(name = "ExcelExport", description = "", category = "", isAllowOverride = false, aspects = {
            "isAsync:false"})
    @ThingworxServiceResult(name = "Result", description = "", baseType = "STRING", aspects = {})
    public String ExcelExport(
            @ThingworxServiceParameter(name = "infotable", description = "", baseType = "INFOTABLE", aspects = {
                    "isRequired:true"}) InfoTable infotable,
            @ThingworxServiceParameter(name = "fileRepository", description = "File repository name", baseType = "THINGNAME", aspects = {
                    "isRequired:true", "thingTemplate:FileRepository"}) String fileRepository,
            @ThingworxServiceParameter(name = "templatePath", description = "templatePath to file", baseType = "STRING", aspects = {
                    "isRequired:true"}) String templatePath,
            @ThingworxServiceParameter(name = "downloadPath", description = "downloadPath to file", baseType = "STRING", aspects = {
                    "isRequired:true"}) String downloadPath
    ) throws Exception {

        long start = System.currentTimeMillis();

        String result = "";

        Thing thing = ThingUtilities.findThing(fileRepository);

        if (thing == null) {
            throw new InvalidRequestException("File Repository [" + fileRepository + "] Does Not Exist", RESTAPIConstants.StatusCode.STATUS_NOT_FOUND);
        } else if (!(thing instanceof FileRepositoryThing)) {
            throw new InvalidRequestException("Thing [" + fileRepository + "] Is Not A File Repository", RESTAPIConstants.StatusCode.STATUS_NOT_FOUND);
        } else {
            FileRepositoryThing repo = (FileRepositoryThing) thing;

            try (InputStream inputStream = repo.openFileForRead(templatePath)) {

                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setWrapText(true);

                XSSFSheet sheet = null;

                InfoTableFunctions infoTableFunctions = new InfoTableFunctions();
                InfoTable distinct = infoTableFunctions.Distinct(infotable, MODULE);

                for (int i = 0; i < distinct.getRowCount(); i++) {

                    String module = distinct.getRow(i).getStringValue(MODULE);
                    InfoTable queryInfo = infoTableFunctions.EQFilter(infotable, MODULE, module);

                    if (module.startsWith("CG")) {
                        sheet = workbook.getSheetAt(0);
                    } else if (module.startsWith("HSG")) {
                        sheet = workbook.getSheetAt(1);
                    } else if (module.startsWith("Pearl")) {
                        sheet = workbook.getSheetAt(2);
                    }

                    Integer rowCount = queryInfo.getRowCount();
                    ArrayList<FieldDefinition> orderedFieldsByOrdinal = queryInfo.getDataShape().getFields().getOrderedFieldsByOrdinal();

                    short lastCellNum = sheet.getRow(0).getLastCellNum();

                    String stationNo = "";
                    int firstRow = head_row_count;
                    int lastRow = head_row_count;
                    int tempRow = head_row_count;

                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (rowIndex < rowCount - 1)
                            sheet.copyRows(rowIndex + head_row_count, rowIndex + head_row_count, rowIndex + head_row_count + 1, new CellCopyPolicy());

                        Row row = sheet.getRow(rowIndex + head_row_count);

                        int cellIndex = 0;

                        ValueCollection valueCollection = queryInfo.getRow(rowIndex);
                        String no = valueCollection.getValue("No").toString();

                        if (rowIndex == 0) {
                            stationNo = no;
                        }

                        if (rowIndex == rowCount - 1) {
                            lastRow = rowIndex + head_row_count;
                        }

                        if (!no.equals(stationNo)) {
                            stationNo = no;
                            lastRow = rowIndex + head_row_count - 1;
                            tempRow = rowIndex + head_row_count;
                        }

                        for (int columnIndex = 0; columnIndex < lastCellNum; columnIndex++) {

                            Cell cell = CellUtil.getCell(row, columnIndex);
                            CellUtil.setCellStyleProperty(cell, CellUtil.WRAP_TEXT, true);

                            if (rowIndex == 0 && orderedFieldsByOrdinal.get(cellIndex).getName().equals(MATERIAL_CAD)) {
                                sheet.getRow(0).getCell(8).setCellValue((String) valueCollection.getPrimitive(orderedFieldsByOrdinal.get(cellIndex).getName()).getValue());
                                cellIndex++;
                            } else if (orderedFieldsByOrdinal.get(cellIndex).getName().equals(MATERIAL_CAD)) {
                                cellIndex++;
                            }

                            if (valueCollection.getValue(orderedFieldsByOrdinal.get(cellIndex).getName()) != null) {
                                switch (orderedFieldsByOrdinal.get(cellIndex).getBaseType()) {
                                    case STRING:
                                        cell.setCellValue((String) valueCollection.getPrimitive(orderedFieldsByOrdinal.get(cellIndex).getName()).getValue());
                                        break;
                                    case INTEGER:
                                        cell.setCellValue((Integer) valueCollection.getPrimitive(orderedFieldsByOrdinal.get(cellIndex).getName()).getValue());
                                        break;
                                    case NUMBER:
                                        cell.setCellValue((Double) valueCollection.getPrimitive(orderedFieldsByOrdinal.get(cellIndex).getName()).getValue());
                                        break;
                                    case BOOLEAN:
                                        cell.setCellValue((Boolean) valueCollection.getPrimitive(orderedFieldsByOrdinal.get(cellIndex).getName()).getValue());
                                        break;
                                    default:
                                        cell.setCellValue((String) valueCollection.getPrimitive(orderedFieldsByOrdinal.get(cellIndex).getName()).getValue());
                                }
                            }
                            mergingCells(sheet, firstRow, lastRow, cellIndex);
                            cellIndex++;
                        }

                        if (firstRow < tempRow) {
                            firstRow = tempRow;
                        }
                        row.setHeight((short) -1);
                    }
                }

                try (OutputStream outputStream = new FileOutputStream(repo.getRootPath() + downloadPath)) {
//                try (OutputStream outputStream = repo.openFileForWrite(repo.getRootPath() + downloadPath, FileRepositoryThing.FileMode.WRITE)) {
                    workbook.write(outputStream);
                    workbook.close();
                } catch (Exception e) {
                    throw new Exception(e.getMessage());
                }

                result = THINGWORX_FILE_REPOSITORIES + fileRepository + downloadPath;

            } catch (Exception e) {
                throw new Exception(e.getMessage());
            }
        }

        long end = System.currentTimeMillis();
        int time = (int) ((end - start) / 1000);
        _logger.warn("{} : {}", "ExcelExport Execution time : ", time);

        return result;
    }

    private void mergingCells(XSSFSheet sheet, int firstRow, int lastRow, int cellIndex) {
        if (firstRow < lastRow) {
            if (cellsList.contains(cellIndex)) {
                sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, cellIndex, cellIndex));
            }
        }

    }
}
