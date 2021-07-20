package com.exceldownload.www.service;

import com.exceldownload.www.comn.ExcelStyle;
import com.exceldownload.www.dto.ExcelDownloadDto;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelDownloadService {

    @Autowired
    private ExcelStyle style;

    public List<ExcelDownloadDto> getData() {
        List<ExcelDownloadDto> result = new ArrayList<>();

        for (int i = 1; i <= 10; i++) {
            ExcelDownloadDto dto = new ExcelDownloadDto();
            dto.setId(Long.valueOf(i));
            dto.setTitle("제목" + i);
            dto.setContent("내용" + i);
            dto.setWriter("작성자" + i);
            dto.setDate(LocalDate.now().plusDays(i));
            dto.setView(i);
            result.add(dto);
        }
        return result;
    }

    public void excelDownload(HttpServletResponse response) throws Exception {
        try {
            String fileName = "엑셀다운로드테스트";
            response.setHeader("Content-Disposition", "attachment; filename=" + new String(fileName.getBytes("euc-kr"), "8859_1") + ".xlsx"); // 안됨
//            response.setHeader("Content-Disposition","attachment; filename=" + new String(fileName.getBytes("UTF-8"), "ISO-8859-1") + ".xlsx"); // 안됨
            response.setHeader("Content-Transfer-Encoding", "binary"); // 전송 데이타의 body를 인코딩한 방법[인코딩 방식]을 표시함

            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet();
            Row row;
            Cell cell;
            int rowNum = 0; // 열 번호

            /**
             * Header
             */
            CellStyle styleHeader = wb.createCellStyle();
            Font fontHeader = wb.createFont();
            style.setHeaderStyle(styleHeader, fontHeader);

            row = sheet.createRow(rowNum++); // 열 생성

            int headerCellNum = 0;
            String[] titles = {"NO", "제목", "내용", "작성자", "날짜", "조회수"};
            for (String title : titles) {
                cell = row.createCell(headerCellNum++); // 셀 생성
                cell.setCellStyle(styleHeader);
                cell.setCellValue(title);
            }

            /**
             * body
             */
            CellStyle styleBodyCenter = wb.createCellStyle();
            Font fontBodyCenter = wb.createFont();
            style.setBodyCenterStyle(styleBodyCenter, fontBodyCenter);

            CellStyle styleBodyNormal = wb.createCellStyle();
            Font fontBodyNormal = wb.createFont();
            style.setBodyNormalStyle(styleBodyNormal, fontBodyNormal);

            List<ExcelDownloadDto> list = getData();
            for (ExcelDownloadDto dto : list) {
                row = sheet.createRow(rowNum++); // 열 생성

                // NO
                cell = row.createCell(0);
                cell.setCellStyle(styleBodyCenter);
                cell.setCellValue(dto.getId());

                // Title
                cell = row.createCell(1);
                cell.setCellStyle(styleBodyNormal);
                cell.setCellValue(dto.getTitle());

                // Content
                cell = row.createCell(2);
                cell.setCellStyle(styleBodyNormal);
                cell.setCellValue(dto.getContent());

                // Writer
                cell = row.createCell(3);
                cell.setCellStyle(styleBodyCenter);
                cell.setCellValue(dto.getWriter());

                // date
                cell = row.createCell(4);
                cell.setCellStyle(styleBodyCenter);
                cell.setCellValue(dto.getDate().toString());

                // view
                cell = row.createCell(5);
                cell.setCellStyle(styleBodyCenter);
                cell.setCellValue(dto.getView());
            }

            /**
             * 열 크기 조절
             */
            short lastColumn = sheet.getRow(sheet.getLastRowNum()).getLastCellNum(); // 열의 개수
            for (int i = 0; i < lastColumn; i++) {
                sheet.autoSizeColumn(i);
                sheet.setColumnWidth(i, (sheet.getColumnWidth(i)) + 1000);
            }

            wb.write(response.getOutputStream());
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            response.getOutputStream().flush();
            response.getOutputStream().close();
        }
    }
}
