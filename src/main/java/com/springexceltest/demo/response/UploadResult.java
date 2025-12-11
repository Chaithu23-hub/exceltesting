package com.springexceltest.demo.response;

import java.util.ArrayList;
import java.util.List;

public class UploadResult {
    public int rowsProcessed = 0;
    public List<Integer> rowsSkipped = new ArrayList<>();
    public List<String> errors = new ArrayList<>();
    public List<String> missingFields =new ArrayList<>();
}
