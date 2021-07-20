package com.exceldownload.www.controller;

import com.exceldownload.www.dto.ExcelDownloadDto;
import com.exceldownload.www.service.ExcelDownloadService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

@RequiredArgsConstructor
@RestController
public class ExcelDownloadController {

    private final ExcelDownloadService service;

    @GetMapping("/data")
    public List<ExcelDownloadDto> data() {
        return service.getData();
    }

    @GetMapping(value = "/excel", produces = "application/text; charset=utf8")
    public void excelDownload(HttpServletResponse response)  throws Exception {
        service.excelDownload(response);
    }

}
