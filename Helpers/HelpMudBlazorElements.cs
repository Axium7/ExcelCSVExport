namespace ExcelCSVExport.Helpers;

public class Element
{
	public int Number { get; set; } = default!;
	public string Sign { get; set; } = default!;
	public string Name { get; set; } = default!;
	public int Position { get; set; } = default!;
	public double Molar { get; set; } = default!;
	public string Group { get; set; } = "Periodic Table";
}

public static class ElementsList
{
	public static List<Element> GetElements()
	{
		return new List<Element>
{
			new Element { Number = 1, Sign = "H", Name = "Hydrogen", Position = 1, Molar = 1.008 },
			new Element { Number = 2, Sign = "He", Name = "Helium", Position = 18, Molar = 4.0026 },
			new Element { Number = 3, Sign = "Li", Name = "Lithium", Position = 1, Molar = 6.94 },
			new Element { Number = 4, Sign = "Be", Name = "Beryllium", Position = 2, Molar = 9.0122 },
			new Element { Number = 5, Sign = "B", Name = "Boron", Position = 13, Molar = 10.81 },
			new Element { Number = 6, Sign = "C", Name = "Carbon", Position = 14, Molar = 12.011 },
			new Element { Number = 7, Sign = "N", Name = "Nitrogen", Position = 15, Molar = 14.007 },
			new Element { Number = 8, Sign = "O", Name = "Oxygen", Position = 16, Molar = 15.999 },
			new Element { Number = 9, Sign = "F", Name = "Fluorine", Position = 17, Molar = 18.998 },
			new Element { Number = 10, Sign = "Ne", Name = "Neon", Position = 18, Molar = 20.180 },
			new Element { Number = 11, Sign = "Na", Name = "Sodium", Position = 1, Molar = 22.990 },
			new Element { Number = 12, Sign = "Mg", Name = "Magnesium", Position = 2, Molar = 24.305 },
			new Element { Number = 13, Sign = "Al", Name = "Aluminium", Position = 13, Molar = 26.982 },
			new Element { Number = 14, Sign = "Si", Name = "Silicon", Position = 14, Molar = 28.085 },
			new Element { Number = 15, Sign = "P", Name = "Phosphorus", Position = 15, Molar = 30.974 },
			new Element { Number = 16, Sign = "S", Name = "Sulfur", Position = 16, Molar = 32.06 },
			new Element { Number = 17, Sign = "Cl", Name = "Chlorine", Position = 17, Molar = 35.45 },
			new Element { Number = 18, Sign = "Ar", Name = "Argon", Position = 18, Molar = 39.948 },
			new Element { Number = 19, Sign = "K", Name = "Potassium", Position = 1, Molar = 39.098 },
			new Element { Number = 20, Sign = "Ca", Name = "Calcium", Position = 2, Molar = 40.078 },
            // Add more elements as needed
        };
	}
}
