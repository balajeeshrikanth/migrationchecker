package com.stonex.gpp.definition;

import java.util.ArrayList;
import java.util.List;

public class ErrorResponse {
    private boolean result;
    private List<AppError> appErrorList;

    public ErrorResponse() {
        this.result = true;
        this.appErrorList = new ArrayList<AppError>();
    }

    public ErrorResponse(boolean result, List<AppError> appErrorList) {
        this.result = result;
        this.appErrorList = appErrorList;
    }

    public boolean isResult() {
        return result;
    }

    public void setResult(boolean result) {
        this.result = result;
    }

    public List<AppError> getAppErrorList() {
        return appErrorList;
    }

    public void setAppErrorList(List<AppError> appErrorList) {
        this.appErrorList = appErrorList;
    }
}
