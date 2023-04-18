package main;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;

public class VBSExecutor {

    public static void main(String[] args) {
        try {
            Path filePath = Paths.get("D:/runMacroScript.vbs");
            Runtime.getRuntime().exec( String.format("wscript %s",filePath));
        }
        catch( IOException e ) {
            e.printStackTrace();
        }
    }

}

