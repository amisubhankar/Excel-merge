package com.excelmerge.Excel_Merge;

import com.excelmerge.Excel_Merge.Config.ReadNWriteExcel;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelMergeApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelMergeApplication.class, args);
		ReadNWriteExcel obj = new ReadNWriteExcel();
		obj.test();
	}

}
