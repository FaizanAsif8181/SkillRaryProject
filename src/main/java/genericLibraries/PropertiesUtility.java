package genericLibraries;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;
/**
 * This class contains reusable methods to perform actions on Properties file
 * @author basheera unnisa
 *
 */

public class PropertiesUtility {
	private Properties property;
	/**
	 * This method is used to initialize Properties file
	 * @param filepath
	 */

	public void propertiesInitialization(String filepath) {
		FileInputStream fis = null;
				
				try {
					fis = new FileInputStream(filepath);
				}catch(FileNotFoundException e) {
					e.printStackTrace();
				}
				property = new Properties();
	
	try {
		property.load(fis);
	}catch(IOException e) {
		e.printStackTrace();
					
				}
		
	}
	/**
	 * This method is used to read data from Properties file
	 * @param key
	 * @return
	 */


public String fetchProperty(String key) {
	return property.getProperty(key);
}
}