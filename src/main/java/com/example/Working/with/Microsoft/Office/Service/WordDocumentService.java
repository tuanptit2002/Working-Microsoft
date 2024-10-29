package com.example.Working.with.Microsoft.Office.Service;

import org.springframework.core.io.ByteArrayResource;

import java.io.IOException;

public interface WordDocumentService {

    public ByteArrayResource createWordDocument() throws IOException;
}
