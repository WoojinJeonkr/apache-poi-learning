package com.example.apachepoilearning.api;

import com.example.apachepoilearning.domain.service.DownloadService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.nio.charset.StandardCharsets;

/**
 * GET 요청을 통해 데이터베이스의 사용자 정보를 포함한 Excel 파일을 다운로드합니다.
 */
@Controller
public class DownloadController {

    // Excel 파일 생성 비즈니스 로직을 담당하는 서비스 의존성 선언
    private final DownloadService downloadService;

    // DownloadService를 주입받기 위한 생성자
    public DownloadController(DownloadService downloadService) {
        this.downloadService = downloadService;
    }

    /**
     * "/download" 경로로 들어오는 GET 요청을 처리하여 Excel 파일을 클라이언트에 제공합니다.
     * @return 생성된 Excel 파일의 바이트 배열과 적절한 HTTP 헤더, HTTP 상태 코드(200 OK)를 포함하는 ResponseEntity
     */
    @GetMapping("/download")
    public ResponseEntity<?> downloadProcess(){

        // 엑셀 데이터는 byte[]로 받음
        byte[] excelBytes = downloadService.downloadXlsx();

        // 클라이언트(브라우저)에게 다운로드될 파일의 원본 이름
        String fileName = "다운로드된_엑셀_파일.xlsx";

        // 1. 파일명을 UTF-8 문자셋으로 URL 인코딩
        // 2. URLEncoder는 공백을 '+'로 인코딩하므로, HTTP 헤더 표준에 맞게 이를 '%20'으로 다시 치환
        String encodedFileName = java.net.URLEncoder.encode(fileName, StandardCharsets.UTF_8).replaceAll("\\+", "%20");

        // 응답 헤더 (엑셀 파일로)
        HttpHeaders headers = new HttpHeaders();
        // Content-Disposition 헤더 설정: 파일을 다운로드하도록 지시하며, 다운로드될 파일명을 지정합니다.
        // "attachment"는 브라우저가 파일을 다운로드하도록 유도하고, "filename"은 파일명을 명시합니다.
        headers.add("Content-Disposition", "attachment; filename=\"" + encodedFileName + "\"");
        // Content-Type 헤더 설정: 응답 본문의 미디어 타입이 Excel .xlsx 파일임을 브라우저에 알립니다.
        headers.add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        // 최종 ResponseEntity 반환:
        // 1. excelBytes: 응답 본문에 포함될 Excel 파일 데이터 (byte 배열)
        // 2. headers: 위에서 설정한 HTTP 응답 헤더
        // 3. HttpStatus.OK: HTTP 상태 코드 200 (성공)
        return new ResponseEntity<>(excelBytes, headers, HttpStatus.OK);
    }

}
