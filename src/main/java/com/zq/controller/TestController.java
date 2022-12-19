package com.zq.controller;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author ZQ
 */
@Controller
@RequestMapping("file")
public class TestController {
    private static final Logger logger = LoggerFactory.getLogger(TestController.class);
    private static final String downloadFilename = "解析结果.xls";
    private static final String[] wordArr = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"};

    @PostMapping("upload")
    public String upload(MultipartFile previousFile, MultipartFile currentFile, MultipartFile compareFile) {
        if (previousFile == null || previousFile.isEmpty() || currentFile == null || currentFile.isEmpty() || compareFile == null || compareFile.isEmpty()) {
            logger.error("file is empty");
            return null;
        }

        FileOutputStream outputStream = null;
        try {
            Workbook previousWorkbook = new HSSFWorkbook(previousFile.getInputStream());
            Workbook currentWorkbook = new HSSFWorkbook(currentFile.getInputStream());
            Workbook compareWorkbook = new HSSFWorkbook(compareFile.getInputStream());

            Map<String, Map<String, Double>> previousMap = new HashMap<>();
            //Map<String, Double> currentMap = new HashMap<>();
            Map<String, List<String>> compareMap = new HashMap<>();

            Sheet compareWorkbookSheet = compareWorkbook.getSheetAt(0);
            for (int i = 0; i <= compareWorkbookSheet.getLastRowNum(); i++) {
                Row row = compareWorkbookSheet.getRow(i);
                String sheetCode = row.getCell(0).getStringCellValue().trim();
                String indexCode = row.getCell(1).getStringCellValue().trim();
                List<String> indexCodeList;
                if (!compareMap.containsKey(sheetCode)) {
                    indexCodeList = new ArrayList<>();
                    compareMap.put(sheetCode, indexCodeList);
                } else {
                    indexCodeList = compareMap.get(sheetCode);
                }
                indexCodeList.add(indexCode);
            }
            //logger.info("待比较内容{}", compareMap);

            resolveExcel(previousWorkbook, previousMap, compareMap);
            //resolveExcel(currentWorkbook, currentMap, compareMap);

            //Map<String, String> colorMap = new HashMap<>(previousMap.size());
            generateColor(currentWorkbook, previousMap);

            File exportFile = new File("D:\\" + downloadFilename);
            outputStream = new FileOutputStream(exportFile);
            currentWorkbook.write(outputStream);
            outputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(outputStream);
        }
        return "redirect:http://localhost:8080";
    }

    private void generateColor(Workbook workbook, Map<String, Map<String, Double>> previousMap) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            // 报表代码
            String sheetCode = sheet.getSheetName();
            // 若比较文档中不包含该报表代码，则不用处理
            Map<String, Double> previousValueMap = previousMap.get(sheetCode);
            if (MapUtils.isEmpty(previousValueMap)) {
                continue;
            }

            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                int cellIndex = 0;
                String num = null;
                for (Cell cell : row) {
                    String stringCellValue;
                    if (cell.getCellType().equals(CellType.STRING) && !StringUtils.isEmpty((stringCellValue = cell.getStringCellValue().trim())) && stringCellValue.contains(".")) {
                        num = stringCellValue.replaceAll("^(\\d[\\d\\\\.]+).+", "$1");
                        if (!StringUtils.isEmpty(num)) {
                            cellIndex = cell.getColumnIndex();
                            break;
                        }
                    }
                }

                if (StringUtils.isEmpty(num)) {
                    continue;
                }

                for (int k = ++cellIndex; k < row.getLastCellNum(); k++) {
                    String index = num + wordArr[k - cellIndex];
                    if (previousValueMap.containsKey(index)) {
                        //logger.info("比对内容：{}", index);
                        Cell cell = row.getCell(k);

                        Double previousValue = previousValueMap.get(index);
                        Double currentValue = cell.getNumericCellValue();
                        /**
                         * 当期和前期进行比较
                         * 比上期变动幅度大于100% 棕黄色
                         * 比上期变动幅度等于100% 黄色
                         * 比上期变动幅度小于-100% 红色
                         * 本期数据为0，上期不等于0 绿色
                         * 本期数据不等于0，上期数据为0 蓝色
                         */
                        CellStyle previousStyle = cell.getCellStyle();
                        CellStyle cellStyle = workbook.createCellStyle();
                        cellStyle.cloneStyleFrom(previousStyle);
                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        cell.setCellStyle(cellStyle);
                        if (previousValue != 0 && previousValue * 2 < currentValue) {
                            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());
                            //logger.info("{}设置棕黄色", index);
                        } else if (previousValue != 0 && previousValue * 2 == currentValue) {
                            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
                            //logger.info("{}设置黄色", index);
                        } else if (currentValue != 0 && previousValue > currentValue * 2) {
                            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
                            //logger.info("{}设置红色", index);
                        } else if (previousValue != 0 && currentValue == 0) {
                            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
                            //logger.info("{}设置绿色", index);
                        } else if (previousValue == 0 && currentValue != 0) {
                            cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
                            //logger.info("{}设置蓝色", index);
                        }
                    }
                }
            }
        }
    }

    private void resolveExcel(Workbook workbook, Map<String, Map<String, Double>> previousMap, Map<String, List<String>> compareMap) {
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            // 报表代码
            String sheetCode = sheet.getSheetName();
            // 若比较文档中不包含该报表代码，则不用处理
            List<String> indexList = compareMap.get(sheetCode);
            if (CollectionUtils.isEmpty(indexList)) {
                continue;
            }
            Map<String, Double> map = new HashMap<>();
            previousMap.put(sheetCode, map);

            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                int cellIndex = 0;
                String num = null;
                for (Cell cell : row) {
                    String stringCellValue;
                    if (cell.getCellType().equals(CellType.STRING) && !StringUtils.isEmpty((stringCellValue = cell.getStringCellValue().trim())) && stringCellValue.contains(".")) {
                        num = stringCellValue.replaceAll("^(\\d[\\d\\\\.]+).+", "$1");
                        if (!StringUtils.isEmpty(num)) {
                            cellIndex = cell.getColumnIndex();
                            break;
                        }
                    }
                }

                if (StringUtils.isEmpty(num)) {
                    continue;
                }

                for (int k = ++cellIndex; k < row.getLastCellNum(); k++) {
                    String index = num + wordArr[k - cellIndex];
                    if (indexList.contains(index)) {
                        map.put(index, row.getCell(k).getNumericCellValue());
                    }
                }
            }
        }
    }

    @GetMapping("downloadReport")
    public void downloadReport(HttpServletResponse response) {
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(downloadFilename.getBytes("gb2312"), "ISO-8859-1"));
            File reportFile = new File("D:\\" + downloadFilename);
            IOUtils.write(FileUtils.readFileToByteArray(reportFile), response.getOutputStream());
            reportFile.delete();
        } catch (Exception e) {
            logger.error(e.toString());
        }
    }
}
