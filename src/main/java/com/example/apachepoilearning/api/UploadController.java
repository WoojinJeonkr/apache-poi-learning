package com.example.apachepoilearning.api;

import com.example.apachepoilearning.domain.download.service.UploadService;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Controller
public class UploadController {

    private final UploadService uploadService;

    public UploadController(UploadService uploadService) {
        this.uploadService = uploadService;
    }

    @GetMapping("/upload")
    public String uploadPage() {
        return "upload";
    }

    @PostMapping("/upload")
    public String uploadProcess(@RequestParam("file") MultipartFile file) throws IOException {
        uploadService.uploadXlsx(file);

        return "redirect:/upload";
    }
}
