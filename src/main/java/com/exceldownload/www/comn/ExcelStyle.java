package com.exceldownload.www.comn;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.stereotype.Component;

@Component
public class ExcelStyle {

    public void setHeaderStyle(CellStyle style, Font font) throws Exception {
        font.setFontName("맑은 고딕"); // 폰트 종류
        font.setFontHeight((short)(9 * 20)); // 폰트 크기
        font.setBold(true);

        style.setFont(font); // 폰트 저장
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex()); // 전경 채우기 색상을 인덱스 색상 값으로 설정
        style.setFillPattern(CellStyle.SOLID_FOREGROUND); // 1로 설정하면 셀이 전경색으로 채워집니다.
        style.setAlignment(CellStyle.ALIGN_CENTER); // 셀의 수평 정렬 유형 설정

        setCommonCellStyle(style);
    }

    public void setBodyCenterStyle(CellStyle style, Font font) throws Exception {
        font.setFontName("맑은 고딕"); // 폰트 종류
        font.setFontHeight((short)(9 * 20)); // 폰트 크기

        style.setFont(font); // 폰트 저장
        style.setAlignment(CellStyle.ALIGN_CENTER); // 셀의 수평 정렬 유형 설정

        setCommonCellStyle(style);
    }

    public void setBodyNormalStyle(CellStyle style, Font font) throws Exception {
        font.setFontName("맑은 고딕"); // 폰트 종류
        font.setFontHeight((short)(9 * 20)); // 폰트 크기

        style.setFont(font); // 폰트 저장

        setCommonCellStyle(style);
    }

    public void setCommonCellStyle(CellStyle style) throws Exception {
        /**
         * 셀 테두리 border
         */
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);

        /**
         * 정렬
         */
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER); // 세로 가운데 정렬
    }

}
