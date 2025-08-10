package com.example.apachepoilearning.api;

import com.example.apachepoilearning.domain.upload.service.UploadService;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

/**
 * 엑셀 파일 업로드 기능 담당
 * 클라이언트의 요청을 받아 파일을 처리하고 응답을 반환합니다.
 */
@Controller
public class UploadController {

    private final UploadService uploadService;

    public UploadController(UploadService uploadService) {
        this.uploadService = uploadService;
    }

    /**
     * GET 요청으로 /upload 경로에 접근했을 때 업로드 페이지를 반환합니다.
     */
    @GetMapping("/upload")
    public String uploadPage() {
        return "upload";
    }

    /**
     * POST 요청으로 /upload 경로에 파일이 전송되었을 때 파일을 처리합니다.
     * 전송된 엑셀 파일을 읽어 비즈니스 로직을 수행합니다.
     *
     * @param file 클라이언트로부터 전송된 MultipartFile 객체 (업로드된 파일 데이터)
     * @return 파일 처리 후 다시 /upload 경로로 리다이렉트합니다.
     * @throws IOException 파일 처리 중 발생할 수 있는 입출력 예외
     */
    @PostMapping("/upload")
    public String uploadProcess(@RequestParam("file") MultipartFile file) throws Exception {
        // uploadService.uploadXlsx(file);
        uploadService.uploadSAXXlsx(file);
        return "redirect:/upload";
    }
}
