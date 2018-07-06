package net.nullpunkt.exceljson.pojo;

import java.util.ArrayList;

public class HeaderObject {

	private String action;
	private String api;
	private Boolean skipstep;
	private String selector;
    private String selectorName;
    private String[] options;
	private String description;
    private String expected;
    private String value;


    // GET/SET

    public String getAction() {
        return action;
    }

    public void setAction(String action) {
        this.action = action;
    }

    public String getApi() {
        return api;
    }

    public void setApi(String api) {
        this.api = api;
    }

    public Boolean getSkipstep() { return skipstep; }

    public void setSkipstep(Boolean skipstep) { this.skipstep = skipstep; }

    public String getSelector() {
        return selector;
    }

    public void setSelector(String selector) {
        this.selector = selector;
    }

    public String getSelectorName() {
        return selectorName;
    }

    public void setSelectorName(String selectorName) {
        this.selectorName = selectorName;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public String getExpected() {
        return expected;
    }

    public void setExpected(String expected) { this.expected = expected; }

    public String[] getOptions() { return options;}

    public void setOptions(String[] options) { this.options = options; }

    public String getValue() { return value; }

    public void setValue(String value) { this.value = value; }

}
