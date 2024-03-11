package com.example.asposecells;

import com.aspose.cells.*;

public class AsposecellsApplication {

	public static void main(String[] args) {
		System.out.println("Application started");

		try (var is = AsposecellsApplication.class.getResourceAsStream("/example_copy_validation_range.xlsx")) {
			Workbook workbook = new Workbook(is);
			Worksheet worksheet = workbook.getWorksheets().get(0);
			Cells cells = worksheet.getCells();

			cells.insertColumns(4, 1);

			System.out.printf("B. text1: '%s', %s, validation2: '%s'%n", cells.get("B1").getValue(), cells.get("B2").getStyle().getForegroundColor(), cells.get("B2").getValidation());
			System.out.printf("C. text1: '%s', %s, validation2: '%s'%n", cells.get("C1").getValue(), cells.get("C2").getStyle().getForegroundColor(), cells.get("C2").getValidation());
			System.out.printf("D. text1: '%s', %s, validation2: '%s'%n", cells.get("D1").getValue(), cells.get("D2").getStyle().getForegroundColor(), cells.get("D2").getValidation());
			System.out.printf("E. text1: '%s', %s, validation2: '%s'%n", cells.get("E1").getValue(), cells.get("E2").getStyle().getForegroundColor(), cells.get("E2").getValidation());
			System.out.printf("F. text1: '%s', %s, validation2: '%s'%n", cells.get("F1").getValue(), cells.get("F2").getStyle().getForegroundColor(), cells.get("F2").getValidation());
			System.out.println("New column created");

			Range range = worksheet.getCells().createRange("B1", "C2");
			Range copyRange = worksheet.getCells().createRange("D1", "E2");

			PasteOptions options = new PasteOptions();
			options.setPasteType(PasteType.ALL);
			copyRange.copy(range);

			System.out.printf("B. text1: '%s', %s, validation2: '%s'%n", cells.get("B1").getValue(), cells.get("B2").getStyle().getForegroundColor(), cells.get("B2").getValidation());
			System.out.printf("C. text1: '%s', %s, validation2: '%s'%n", cells.get("C1").getValue(), cells.get("C2").getStyle().getForegroundColor(), cells.get("C2").getValidation());
			System.out.printf("D. text1: '%s', %s, validation2: '%s'%n", cells.get("D1").getValue(), cells.get("D2").getStyle().getForegroundColor(), cells.get("D2").getValidation());
			System.out.printf("E. text1: '%s', %s, validation2: '%s'%n", cells.get("E1").getValue(), cells.get("E2").getStyle().getForegroundColor(), cells.get("E2").getValidation());
			System.out.printf("F. text1: '%s', %s, validation2: '%s'%n", cells.get("F1").getValue(), cells.get("F2").getStyle().getForegroundColor(), cells.get("F2").getValidation());
			System.out.println("Copying completed");
			workbook.save("result.xlsx");
			System.out.println("Result saved");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

}
