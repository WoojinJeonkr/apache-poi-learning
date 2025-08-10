package com.example.apachepoilearning.domain.download.service;

import com.example.apachepoilearning.entity.User;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class DownloadService {

    public byte[] downloadXlsx() throws IOException {

        // 엑셀 Workbook 객체 생성
        // XSSFWorkbook은 엑셀 파일을 읽을 때 모든 데이터를 메모리에 로드
        // 데이터가 많을수록 더 많은 메모리를 사용하게 되고,
        // 시스템에서 할당된 메모리보다 더 많은 메모리를 요구하면 OOM(OutOfMemoryError) 에러가 발생
        Workbook workbook = new XSSFWorkbook();

        // 시트 추가
        // ------------------------- Sheet1: 사용자 목록 데이터 -------------------------
        Sheet sheet1 = workbook.createSheet("User List"); // "User List"라는 이름의 첫 번째 시트 생성

        // 사용자 데이터 생성 (하드코딩 대신 동적 생성 예시)
        List<User> users = new ArrayList<>();
        users.add(new User(1L, "김철수", "kim@example.com", 25));
        users.add(new User(2L, "이영희", "lee@example.com", 30));
        users.add(new User(3L, "박민준", "park@example.com", 28));
        users.add(new User(4L, "최수정", "choi@example.com", 32));

        // Sheet1 헤더 로우 생성
        Row headerRow1 = sheet1.createRow(0);
        String[] headers1 = {"ID", "이름", "이메일", "나이"};
        for (int i = 0; i < headers1.length; i++) {
            headerRow1.createCell(i).setCellValue(headers1[i]);
        }

        // Sheet1 데이터 로우 생성
        int rowNum1 = 1;
        for (User user : users) {
            Row row = sheet1.createRow(rowNum1++);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getName());
            row.createCell(2).setCellValue(user.getEmail());
            row.createCell(3).setCellValue(user.getAge());
        }

        // ------------------------- Sheet2: 통계 요약 데이터 -------------------------
        Sheet sheet2 = workbook.createSheet("Statistics Summary"); // "Statistics Summary"라는 이름의 두 번째 시트 생성

        // 통계 데이터 생성 (예시)
        Map<String, Integer> stats = new LinkedHashMap<>(); // 순서 유지를 위해 LinkedHashMap 사용
        stats.put("Category A", 150);
        stats.put("Category B", 230);
        stats.put("Category C", 90);
        stats.put("Category D", 400);

        // Sheet2 헤더 로우 생성
        Row headerRow2 = sheet2.createRow(0);
        headerRow2.createCell(0).setCellValue("Category");
        headerRow2.createCell(1).setCellValue("Count");

        // Sheet2 데이터 로우 생성
        int rowNum2 = 1;
        for (Map.Entry<String, Integer> entry : stats.entrySet()) {
            Row row = sheet2.createRow(rowNum2++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getValue());
        }

        // ------------------------- Sheet3: 날짜 및 숫자 데이터 -------------------------
        Sheet sheet3 = workbook.createSheet("Daily Data"); // "Daily Data"라는 이름의 세 번째 시트 생성

        // 날짜 데이터 스타일 생성 (날짜 형식 지정)
        CellStyle dateCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd"));

        // Sheet3 헤더 로우 생성
        Row headerRow3 = sheet3.createRow(0);
        headerRow3.createCell(0).setCellValue("날짜");
        headerRow3.createCell(1).setCellValue("값1");
        headerRow3.createCell(2).setCellValue("값2");

        // Sheet3 데이터 로우 생성 (지난 7일간의 데이터 예시)
        int rowNum3 = 1;
        for (int i = 0; i < 7; i++) {
            Row row = sheet3.createRow(rowNum3++);
            LocalDate date = LocalDate.now().minusDays(i); // 오늘부터 i일 전의 날짜
            double value1 = 100 + (i * 5) + Math.random() * 10; // 임의의 숫자 값
            double value2 = 50 - (i * 2) + Math.random() * 5;  // 임의의 숫자 값

            // 날짜 셀
            Cell dateCell = row.createCell(0);
            dateCell.setCellValue(date); // LocalDate 객체를 직접 셀에 설정
            dateCell.setCellStyle(dateCellStyle); // 날짜 스타일 적용

            // 숫자 셀
            row.createCell(1).setCellValue(value1);
            row.createCell(2).setCellValue(value2);
        }

        // ------------------------- Sheet3: 셀 기능 조작 연습 -------------------------
        Sheet sheet4 = workbook.createSheet("Cell Practice");

        // 셀 병합: B1 (0행, 1열)부터 C1 (0행, 2열)까지 병합
        sheet4.addMergedRegion(new CellRangeAddress(0, 0, 1, 2));

        // 1. 셀 디자인
        // 스타일 객체 생성
        CellStyle style1 = workbook.createCellStyle();

        // 배경색
        style1.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 테두리 스타일
        style1.setBorderTop(BorderStyle.THICK);
        style1.setBorderBottom(BorderStyle.THICK);
        style1.setBorderLeft(BorderStyle.THICK);
        style1.setBorderRight(BorderStyle.THICK);

        // 테두리 색상
        style1.setTopBorderColor(IndexedColors.RED.getIndex());
        style1.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        style1.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style1.setRightBorderColor(IndexedColors.PINK.getIndex());

        // 특정 셀에 저장 (여러 셀이라면 반복문을 통해 반복 수행)
        Row sheet4_row1 = sheet4.createRow(1);
        Cell sheet4_row1_cell0 = sheet4_row1.createCell(0);
        sheet4_row1_cell0.setCellValue("데이터");
        sheet4_row1_cell0.setCellStyle(style1);

        // 2. 글자 디자인
        CellStyle style2 = workbook.createCellStyle();

        // 폰트 설정
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short)14);     // 폰트 크기
        font.setBold(true);                        // 굵게
        font.setItalic(true);                      // 이탤릭
        font.setColor(IndexedColors.BLUE.getIndex()); // 폰트 색상
        style2.setFont(font);

        // 글자 정렬
        style2.setAlignment(HorizontalAlignment.CENTER); // 가로 가운데 정렬
        style2.setVerticalAlignment(VerticalAlignment.CENTER); // 세로 가운데 정렬

        // 특정 셀에 지정
        Row sheet4_row2 = sheet4.createRow(2);
        Cell sheet4_row2_cell0 = sheet4_row2.createCell(0);
        sheet4_row2_cell0.setCellValue("데이터");
        sheet4_row2_cell0.setCellStyle(style1);

        // 3. 필터: 0번째 행의 0열부터 2열까지 필터 적용 (A1:C1)
        sheet4.setAutoFilter(new CellRangeAddress(0, 0, 0, 2));

        // 출력 스트림화
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();

        return outputStream.toByteArray();
    }

}
