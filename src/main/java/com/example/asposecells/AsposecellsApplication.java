package com.example.asposecells;

import com.aspose.cells.*;

public class AsposecellsApplication {

	public static void main(String[] args) {
		System.out.println("Application started");

		try (var is = AsposecellsApplication.class.getResourceAsStream("/example_copy_validation_range.xlsx")) {
			Workbook workbook = new Workbook(is);
			Worksheet worksheet = workbook.getWorksheets().get(0);

			/*Range range = worksheet.getCells().createRange(0, 1, 1, 2);
			Range copyRange = worksheet.getCells().createRange(0, 3, 1, 2);

			PasteOptions options = new PasteOptions();
			options.setPasteType(PasteType.ALL);
			copyRange.copy(range);*/

			worksheet.getCells().insertColumns(2, 2);

			System.out.printf("B1. text: '%s', validation: '%s'%n", worksheet.getCells().get("B1").getValue(), worksheet.getCells().get("B2").getValidation());
			System.out.printf("C1. text: '%s', validation: '%s'%n", worksheet.getCells().get("C1").getValue(), worksheet.getCells().get("C2").getValidation());
			System.out.printf("D1. text: '%s', validation: '%s'%n", worksheet.getCells().get("D1").getValue(), worksheet.getCells().get("D2").getValidation());
			System.out.printf("E1. text: '%s', validation: '%s'%n", worksheet.getCells().get("E1").getValue(), worksheet.getCells().get("E2").getValidation());
			System.out.printf("F1. text: '%s', validation: '%s'%n", worksheet.getCells().get("F1").getValue(), worksheet.getCells().get("F2").getValidation());
			System.out.println("Copying completed");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

}
