package com.zq.controller;

import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.NumberFormat;
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
    private static final String COMPARE_FILE = "校验文件.xls";
    private static final String COMPARE_TEMPLATE_FILE = "compare_template.xls";
    private static final String RESULT_FILE = "解析结果.xls";
    private static final String RESULT_FILE_SUFFIX = "_result.xls";
    static final String BASE_PATH = System.getProperty("user.dir") + File.separator;
    static final String[] titles = {"报表代码", "指标位置", "指标代码", "指标名称", "本期数据值", "上期数据值", "比上期（万元）", "比上期（%）", "备注"};

    @PostMapping("upload")
    @ResponseBody
    @CrossOrigin("*")
    public String upload(MultipartFile[] previousFile, MultipartFile[] currentFile, HttpServletRequest request) {
        if (previousFile == null || previousFile.length == 0) {
            logger.error("请上传上期文件夹");
            return "请上传上期文件夹";
        }
        if (currentFile == null || currentFile.length == 0) {
            logger.error("请上传当期文件夹");
            return "请上传当期文件夹";
        }

        // 判断比对文件是否存在
        String ipAddr = getIpAddr(request);
        File compareFile = new File(BASE_PATH + ipAddr + ".xls");
        if (compareFile == null || !compareFile.exists()) {
            logger.error("请上传校验文件");
            return "请上传校验文件";
        }

        FileOutputStream outputStream = null;
        try {
            ArrayListValuedHashMap<String, String> locationMap = new ArrayListValuedHashMap<>();
            Map<String, List<Double>> previousMap = new HashMap<>();
            Map<String, List<Double>> currentMap = new HashMap<>();

            // 先解析校验文件
            Map<String, List<String>> compareMap = doHandlerCompareFile(compareFile, locationMap);

            // 解析文件夹
            doHandlerDirectory(previousFile, previousMap, compareMap);
            doHandlerDirectory(currentFile, currentMap, compareMap);

            // 开始比对
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet("数据比对结果");
            HSSFRow firstRow = sheet.createRow(0);
            for (int i = 0; i < titles.length; i++) {
                HSSFCell cell = firstRow.createCell(i);
                cell.setCellValue(titles[i]);
            }

            int i = 1;
            for (Map.Entry<String, List<String>> entry : compareMap.entrySet()) {
                String sheetCode = entry.getKey();
                List<Double> previousValueList = previousMap.get(sheetCode);
                List<Double> currentValueList = currentMap.get(sheetCode);
                for (int j = 0; j < entry.getValue().size(); j++) {
                    Double previousValue = previousValueList.get(j);
                    Double currentValue = currentValueList.get(j);
                    /**
                     * 当期和前期进行比较
                     * 比上期变动幅度大于100% 棕黄色
                     * 比上期变动幅度等于100% 黄色
                     * 比上期变动幅度小于-100% 红色
                     * 本期数据为0，上期不等于0 绿色
                     * 本期数据不等于0，上期数据为0 蓝色
                     */
                    String remarks;
                    CellStyle cellStyle = wb.createCellStyle();
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    if (previousValue != 0 && previousValue * 2 < currentValue) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());
                        remarks = "比上期变动幅度大于100%";
                        //logger.info("{}设置棕黄色", index);
                    } else if (previousValue != 0 && previousValue * 2 == currentValue) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.YELLOW.getIndex());
                        remarks = "比上期变动幅度等于100%";
                        //logger.info("{}设置黄色", index);
                    } else if (currentValue != 0 && previousValue > currentValue * 2) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
                        remarks = "比上期变动幅度小于-100%";
                        //logger.info("{}设置红色", index);
                    } else if (previousValue != 0 && currentValue == 0) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
                        remarks = "本期数据为0，上期不等于0";
                        //logger.info("{}设置绿色", index);
                    } else if (previousValue == 0 && currentValue != 0) {
                        cellStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
                        remarks = "本期数据不等于0，上期数据为0";
                        //logger.info("{}设置蓝色", index);
                    } else {
                        continue;
                    }

                    HSSFRow row = sheet.createRow(i++);
                    int columnIndex = 0;
                    HSSFCell cell0 = row.createCell(columnIndex++);
                    cell0.setCellValue(sheetCode);
                    HSSFCell cell1 = row.createCell(columnIndex++);
                    String indexLocation = entry.getValue().get(j);
                    cell1.setCellValue(indexLocation);
                    HSSFCell cellIndexCode = row.createCell(columnIndex++);
                    List<String> strings = locationMap.get(indexLocation);
                    cellIndexCode.setCellValue(strings.get(0));
                    HSSFCell cellIndexName = row.createCell(columnIndex++);
                    cellIndexName.setCellValue(strings.get(1));
                    HSSFCell cell2 = row.createCell(columnIndex++, CellType.NUMERIC);
                    cell2.setCellValue(currentValue);
                    HSSFCell cell3 = row.createCell(columnIndex++, CellType.NUMERIC);
                    cell3.setCellValue(previousValue);
                    HSSFCell cell4 = row.createCell(columnIndex++);
                    cell4.setCellValue((currentValue - previousValue) / 10000);
                    HSSFCell cell5 = row.createCell(columnIndex++);
                    NumberFormat numberFormat = NumberFormat.getInstance();
                    numberFormat.setMinimumFractionDigits(2);
                    cell5.setCellValue(numberFormat.format(currentValue / previousValue) + "%");
                    HSSFCell cell6 = row.createCell(columnIndex++);
                    cell6.setCellValue(remarks);

                    for (Cell cell : row) {
                        cell.setCellStyle(cellStyle);
                    }
                }
            }

            String resultFileName = BASE_PATH + ipAddr + RESULT_FILE_SUFFIX;
            File resultFile = new File(resultFileName);
            wb.write(resultFile);
        } catch (Exception e) {
            logger.error("", e);
            return "" + e.getMessage();
        } finally {
            IOUtils.closeQuietly(outputStream);
        }
        return "0";
    }

    private String getIpAddr(HttpServletRequest request) {
        String ipAddr = request.getRemoteHost();
        return ipAddr;
        //return "compare";
    }

    private Map<String, List<String>> doHandlerCompareFile(File file, ArrayListValuedHashMap<String, String> locationMap) throws IOException {
        Workbook compareWorkbook = new HSSFWorkbook(new FileInputStream(file));
        Map<String, List<String>> compareMap = new HashMap<>();
        Sheet compareWorkbookSheet = compareWorkbook.getSheetAt(0);
        for (int i = 1; i <= compareWorkbookSheet.getLastRowNum(); i++) {
            Row row = compareWorkbookSheet.getRow(i);
            if (row.getLastCellNum() < 4) {
                throw new RuntimeException("请填写完整第" + (i + 1) + "行校验内容");
            }
            String sheetCode = row.getCell(0).getStringCellValue().trim();
            String indexLocation = row.getCell(1).getStringCellValue().trim(); //位置代码
            String indexCode = row.getCell(2).getStringCellValue().trim(); //指标代码
            String indexName = row.getCell(3).getStringCellValue().trim(); //指标名称
            List<String> indexLocationList;
            if (compareMap.containsKey(sheetCode)) {
                indexLocationList = compareMap.get(sheetCode);
            } else {
                indexLocationList = new ArrayList<>();
                compareMap.put(sheetCode, indexLocationList);
            }
            locationMap.put(indexLocation, indexCode);
            locationMap.put(indexLocation, indexName);
            indexLocationList.add(indexLocation);
        }
        return compareMap;
    }

    private void doHandlerDirectory(MultipartFile[] files, Map<String, List<Double>> valueMap, Map<String, List<String>> compareMap) throws IOException {
        for (MultipartFile listFile : files) {
            String sheetCode = FilenameUtils.getBaseName(listFile.getOriginalFilename());
            // 若比较文档中不包含该报表代码，则不用处理
            List<String> indexList = compareMap.get(sheetCode);
            if (indexList == null || indexList.isEmpty()) {
                continue;
            }
            Workbook workbook = new HSSFWorkbook(listFile.getInputStream());
            Sheet sheet = workbook.getSheetAt(0);
            List<Double> valueList = new ArrayList<>();
            valueMap.put(sheetCode, valueList);

            for (String index : indexList) {
                char ch = index.charAt(0);
                int rowIndex = Integer.valueOf(index.substring(1)) - 1;
                int cellIndex = ch - 'A';
                valueList.add(sheet.getRow(rowIndex).getCell(cellIndex).getNumericCellValue());
            }
        }
    }

    @GetMapping("downloadReport")
    public void downloadReport(HttpServletRequest request, HttpServletResponse response) {
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(RESULT_FILE.getBytes("gb2312"), "ISO-8859-1"));
            File reportFile = new File(BASE_PATH + getIpAddr(request) + RESULT_FILE_SUFFIX);
            IOUtils.write(FileUtils.readFileToByteArray(reportFile), response.getOutputStream());
        } catch (Exception e) {
            logger.error(e.toString());
        }
    }

    @PostMapping("uploadCompareFile")
    @ResponseBody
    @CrossOrigin("*")
    public String upload(MultipartFile file, HttpServletRequest request) {
        FileOutputStream fileOutputStream = null;
        try {
            File reportFile = new File(BASE_PATH + getIpAddr(request) + ".xls");
            fileOutputStream = new FileOutputStream(reportFile);
            IOUtils.write(IOUtils.toByteArray(file.getInputStream()), fileOutputStream);
            fileOutputStream.flush();
            return "0";
        } catch (Exception e) {
            logger.error("", e);
            return e.getMessage();
        } finally {
            IOUtils.closeQuietly(fileOutputStream);
        }
    }

    @GetMapping("downloadCompareFile")
    public void downloadCompareFile(HttpServletRequest request, HttpServletResponse response) {
        try {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(COMPARE_FILE.getBytes("gb2312"), "ISO-8859-1"));
            File reportFile = new File(BASE_PATH + getIpAddr(request) + ".xls");
            if (!reportFile.exists()) {
                reportFile = new File(BASE_PATH + COMPARE_TEMPLATE_FILE);
            }
            IOUtils.write(FileUtils.readFileToByteArray(reportFile), response.getOutputStream());
            response.getOutputStream().flush();
        } catch (Exception e) {
            logger.error("", e);
        }
    }
}
