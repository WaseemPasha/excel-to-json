package net.nullpunkt.exceljson;

import net.nullpunkt.exceljson.convert.ExcelToJsonConverter;
import net.nullpunkt.exceljson.convert.ExcelToJsonConverterConfig;
import net.nullpunkt.exceljson.pojo.ExcelWorkbook;

import net.nullpunkt.exceljson.pojo.ExcelWorksheet;
import org.apache.commons.cli.BasicParser;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Option;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.json.JSONObject;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.Collection;

/**
 * Converting Excel to Json for TestData
 *
 */
public class App {
	
	public static void main( String[] args ) throws Exception
    {
		Options options = new Options();
        options.addOption("s", "source", true, "The source file which should be converted into json.");
        options.addOption("d", "destination", true, "The destination file where json files should be created in the folder.");
		options.addOption("df", "dateFormat", true, "The template to use for fomatting dates into strings.");
		options.addOption("?", "help", true, "This help text.");
		options.addOption(new Option("percent", "Parse percent values as floats."));
		options.addOption(new Option("empty", "Include rows with no data in it."));
		options.addOption(new Option("pretty", "To render output as pretty formatted json."));
		options.addOption(new Option("fillColumns", "To fill rows with null values until they all have the same size."));
		
		CommandLineParser parser = new BasicParser();
		CommandLine cmd = null;
		try {
			cmd = parser.parse(options, args);
		} catch(ParseException e) {
			help(options);
			return; 
		}
		
		if(cmd.hasOption("?")) {
			help(options);
			return;
		}
		
		ExcelToJsonConverterConfig config = ExcelToJsonConverterConfig.create(cmd);
		String valid = config.valid();
		if(valid!=null) {
			System.out.println(valid);
			help(options);
			return;
		}

        // Create a new FileWriter object
        Collection<String> jsonData = new ArrayList<>();
        ExcelWorkbook book = ExcelToJsonConverter.convert(config);
        Collection<ExcelWorksheet> sheets = book.getSheets();

        /*
        * Creating New Directory
        * */
        Path directory = Paths.get(config.getDestinationFile());
        try {
            Files.createDirectories(directory);

        } catch (IOException e) {
            e.printStackTrace();
        }

        /*
        * Converting sheets to JsonObject and creating their files
        * */
        sheets.forEach(sheet ->{
            String sheetJson = sheet.toJson(config.isPretty());
//            jsonData.add(sheetJson);
                /*
                * Creating json files
                * */
                Path jsonFile = Paths.get(config.getDestinationFile()+"/"+sheet.getName()+".json");
                try {
                    Files.createFile(jsonFile);
                    try {
                        Files.write(Paths.get(config.getDestinationFile()+"/"+sheet.getName()+".json"), sheetJson.getBytes(), StandardOpenOption.APPEND);
                    }catch (IOException e) {
                        //exception handling left as an exercise for the reader
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
        });

        jsonData.forEach(json ->{
                System.out.println(json);


        });

		String json = book.toJson(config.isPretty());
		System.out.println("SUCCESSFULLY CONVERTED!!");
    }


	private static void help(Options options) {
		HelpFormatter formater = new HelpFormatter();
		formater.printHelp("java -jar excel-to-json.jar", options);
	}
}
