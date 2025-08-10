package com.example.apachepoilearning.domain.download.service;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Service
public class UploadService {

    public void uploadXlsx(MultipartFile file) throws IOException {

        // 파일 이름 출력
        System.out.println(file.getOriginalFilename());
    }
}
