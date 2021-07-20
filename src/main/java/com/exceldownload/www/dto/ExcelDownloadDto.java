package com.exceldownload.www.dto;

import lombok.Getter;
import lombok.Setter;

import java.time.LocalDate;

@Getter
@Setter
public class ExcelDownloadDto {

    private Long id;

    private String title;

    private String content;

    private String writer;

    private LocalDate date;

    private int view;

}
