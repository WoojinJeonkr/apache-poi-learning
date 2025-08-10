package com.example.apachepoilearning.domain.download.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;

@Service
public class UploadService {

    public void uploadXlsx(MultipartFile file) throws IOException {

        // 파일 이름 출력
        System.out.println(file.getOriginalFilename());

        // 업로드한 파일을 InputStream 형태로 받음
        InputStream inputStream = file.getInputStream();

        // 엑셀 파일 객체 읽기
        Workbook workbook = new XSSFWorkbook(inputStream);

        // 시트 정보
        int sheetCount = workbook.getNumberOfSheets(); // 시트 개수
        for (int i = 0; i < workbook.getNumberOfSheets(); ++i) { // 시트 이름 출력
            Sheet sheet = workbook.getSheetAt(i);
            System.out.println(sheet.getSheetName());
        }

        // 시트 읽기: 특정 시트 선택 (시트 개수만큼 반복문으로 개수 만큼 돌려서 처리하는 케이스 많음)
        Sheet sheet = workbook.getSheetAt(0);

        // 시트별 데이터 처리
        for (Row row : sheet) {
            for (Cell cell : row) {
                System.out.println(cell);
            }
            System.out.println();
        }
    }
}
