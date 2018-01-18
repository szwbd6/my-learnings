package HackerRank;

import java.io.*;
import java.util.*;
import java.text.*;
import java.math.*;
import java.util.regex.*;

public class Day6 {

    public static void main(String[] args) {
        Scanner scan = new Scanner (System.in);
        
        int iInputCount=0;
        String sInput="";

        iInputCount = scan.nextInt();
        scan.nextLine();
        for (int iLoopCounter=0; iLoopCounter<iInputCount; iLoopCounter++)
        {
        	sInput = scan.nextLine();
        	for (int iOddCounter=0; iOddCounter<sInput.length(); iOddCounter += 2)
        	{
        		System.out.print(sInput.charAt(iOddCounter));
        	}
        	System.out.print(" ");
        	for (int iEvenCounter=1; iEvenCounter<sInput.length(); iEvenCounter += 2)
        	{
        		System.out.print(sInput.charAt(iEvenCounter));
        	}  	
        }
        scan.close();
    }
}

