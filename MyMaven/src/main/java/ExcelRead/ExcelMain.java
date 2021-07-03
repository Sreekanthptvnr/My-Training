package ExcelRead;

import java.io.IOException;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		String h=ExcelCode.ReadStringData(0, 0);
		System.out.println(h);
	    String j=ExcelCode.ReadIntegerData(0, 1);
	    System.out.println(j);
	}

}
