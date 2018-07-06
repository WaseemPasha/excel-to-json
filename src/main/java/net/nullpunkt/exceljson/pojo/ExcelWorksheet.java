package net.nullpunkt.exceljson.pojo;

import org.codehaus.jackson.JsonGenerationException;
import org.codehaus.jackson.map.JsonMappingException;
import org.codehaus.jackson.map.ObjectMapper;

import java.io.IOException;

public class ExcelWorksheet {

	private String name;

	private StepObject data = new StepObject();

	public String toJson(boolean pretty) {
		try {
			if(pretty) {
				return new ObjectMapper().writer().withDefaultPrettyPrinter().writeValueAsString(this);
			} else {
				return new ObjectMapper().writer().withPrettyPrinter(null).writeValueAsString(this);
			}
		} catch (JsonGenerationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (JsonMappingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}
	// GET/SET
	
	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public StepObject getData() {
		return data;
	}

	public void setData(StepObject data) {
		this.data = data;
	}
}
