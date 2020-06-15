package main.java.com.docConverter;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import ooo.connector.BootstrapConnector;

import java.io.File;

public class Main {

    public static void main(String[] args) throws Exception, BootstrapException {

        XComponentContext xContext = getxComponentContext();
        XDesktop   xDesktop = getxDesktop(xContext);

        String pathFolder = "/Users/{user_name}/Documents/";//directory to the file you wanted to convert
        String fileToCovert = "convertMe.odt";//file you wanted to convert

        if (!new File(pathFolder + fileToCovert).canRead()) {
            throw new RuntimeException("Cannot load template:" + new File(pathFolder + fileToCovert));
        }

        XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class, xDesktop);

        String inputFile = "file:///"+ pathFolder + fileToCovert;
        PropertyValue propertyValues[] = new PropertyValue[1];
        propertyValues[0] = new PropertyValue();
        propertyValues[0].Name = "Hidden";
        propertyValues[0].Value = Boolean.TRUE;

        // Save the document
        Object loadFileFromURL = xComponentLoader.loadComponentFromURL(inputFile, "_blank", 0,  propertyValues);
        XStorable xStorable = UnoRuntime .queryInterface(XStorable.class, loadFileFromURL);
        propertyValues = new PropertyValue[3];

        // Overwriting
        propertyValues[0] = new PropertyValue();
        propertyValues[0].Name = "Overwrite";
        propertyValues[0].Value = Boolean.TRUE;

        // export odt to pdf
        propertyValues[0] = new PropertyValue();
        propertyValues[0].Name = "FilterName";
        propertyValues[0].Value = "writer_pdf_Export";

        // PDF tagged
        PropertyValue[] propertyData = new PropertyValue[1];
        propertyData[0] = new PropertyValue();
        propertyData[0].Name = "UseTaggedPDF";
        propertyData[0].Value = Boolean.TRUE;
        propertyValues[2] = new PropertyValue();
        propertyValues[2].Name = "FilterData";
        propertyValues[2].Value = propertyData;


        // the url where the pdf is to be saved
        String outputFile = pathFolder + "sample0.pdf";

        xStorable.storeToURL("file:///" + outputFile, propertyValues);

        // shutdown
        xDesktop.terminate();
        System.out.println("Conversion process successfully completed ...");

    }

    private static XDesktop getxDesktop(XComponentContext xContext) throws Exception {
        XMultiComponentFactory xMCF = xContext.getServiceManager();

        String available = (xMCF != null ? "available" : "not available");
        System.out.println( "Remote ServiceManager is " + available );

        Object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);

        return (XDesktop) UnoRuntime.queryInterface(XDesktop.class, oDesktop);
    }

    private static XComponentContext getxComponentContext() throws BootstrapException {
        String hostAndPort = "host=localhost,port=8100";

        // accept option
        String oooAcceptOption = "--accept=socket,"+hostAndPort+";urp;";

        // connection string
        String unoConnectString = "uno:socket,"+hostAndPort+";urp;StarOffice.ComponentContext";

        String libreOfficeFilePath = "/Users/kevinaton/Desktop/LibreOffice.app/Contents/MacOS/soffice";
        BootstrapConnector bootstrapConnector = new BootstrapConnector(libreOfficeFilePath);

        XComponentContext  xContext = bootstrapConnector.connect(oooAcceptOption, unoConnectString);

        if (xContext == null) {
            throw new BootstrapException("no local component context!");
        }
        return xContext;
    }


}
