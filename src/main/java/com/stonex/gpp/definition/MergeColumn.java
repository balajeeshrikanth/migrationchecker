package com.stonex.gpp.definition;

public class MergeColumn {
    private String column;
    private boolean isPrefixed;

    public MergeColumn() {
    }

    public MergeColumn(String column, boolean isPrefixed) {
        this.column = column;
        this.isPrefixed = isPrefixed;
    }

    public String getColumn() {
        return column;
    }

    public void setColumn(String column) {
        this.column = column;
    }

    public boolean isPrefixed() {
        return isPrefixed;
    }

    public void setPrefixed(boolean prefixed) {
        isPrefixed = prefixed;
    }
}
