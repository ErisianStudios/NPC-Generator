using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace NPCGenerator {
    class Program {
        #region Properties
        static string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        //Randomizer.
        static Random rnd = new Random();
        //Genders. Feel free to add more, if that's your thing.
        static string[] gender = { "male", "female" };
        //Male names.
        static string maleNamePath = dir + "\\MaleNames.txt";
        static string[] maleNames = File.ReadAllLines(maleNamePath);
        //Female names.
        static string femaleNamePath = dir + @"\FemaleNames.txt";
        static string[] femaleNames = File.ReadAllLines(femaleNamePath);
        //Surnames
        static string surnamePath = dir + @"\Surnames.txt";
        static string[] surnames = File.ReadAllLines(surnamePath);
        //Fantasy races. Really anything can go in here. could also change it to 'nationality' if not using a fantasy setting.
        static string racesPath = dir + @"\Races.txt";
        static string[] races = File.ReadAllLines(racesPath);
        //Class
        static string classesPath = dir + @"\Classes.txt";
        static string[] classes = File.ReadAllLines(classesPath);
        //Professions
        static string professionsPath = dir + @"\Professions.txt";
        static string[] professions = File.ReadAllLines(professionsPath);
        //Personality traits
        static string traitsPath = dir + @"\Traits.txt";
        static string[] traits = File.ReadAllLines(traitsPath);
        //D&D Classes
        static string DNDClassPath = dir + @"\DNDClasses.txt";
        static string[] DNDClasses = File.ReadAllLines(DNDClassPath);
        static List<NPC> npcs = new List<NPC>();
        #endregion


        static void Main(string[] args) {  
            AskForGeneration();
            GenerateSpreadsheet();
            Console.ReadKey();
        }

        static void GenerateSpreadsheet() {
            string spreadsheetPath = "Output.xlsx";
            File.Delete(spreadsheetPath);
            FileInfo nfo = new FileInfo(spreadsheetPath);
            ExcelPackage pkg = new ExcelPackage(nfo);
            var worksheet = pkg.Workbook.Worksheets.Add("NPCs");
            worksheet.Cells["A1"].Value = "Names";
            worksheet.Cells["B1"].Value = "Surnames";
            worksheet.Cells["C1"].Value = "Gender";
            worksheet.Cells["D1"].Value = "Class";
            worksheet.Cells["E1"].Value = "Race";
            worksheet.Cells["F1"].Value = "Profession";
            worksheet.Cells["G1"].Value = "Trait";
            worksheet.Cells["A1:G1"].Style.Font.Bold = true;

            //Populate spreadsheet.
            int currentRow = 2;
            foreach (var npc in npcs) {
                worksheet.Cells["A" + currentRow.ToString()].Value = npc.Name;
                worksheet.Cells["B" + currentRow.ToString()].Value = npc.Surname;
                worksheet.Cells["C" + currentRow.ToString()].Value = npc.Gender;
                worksheet.Cells["D" + currentRow.ToString()].Value = npc.Class;
                worksheet.Cells["E" + currentRow.ToString()].Value = npc.Race;
                worksheet.Cells["F" + currentRow.ToString()].Value = npc.Profession;
                worksheet.Cells["G" + currentRow.ToString()].Value = npc.Trait;

                currentRow++;
            }
            worksheet.View.FreezePanes(2, 2);
            pkg.Save();
            Console.WriteLine("\n******************************************************************************************\n\nSpreadsheet generated as Output.xslx. Please back this file up, as the next time you run this program it will be replaced.\n\n******************************************************************************************\n");
        }

        static void AskForGeneration() {
            Console.WriteLine("Please enter a number of NPCs to generate.");
            if (int.TryParse(Console.ReadLine(), out int numberToGenerate)) {
                for (int i = 0; i < numberToGenerate; i++) {
                    GenerateNPC();
                }
            } else AskForGeneration();
        }

        static void GenerateNPC() {
            NPC npc = new NPC();
            int gIndex = rnd.Next(gender.Length);
            npc.Gender = gender[gIndex];
            int rIndex = rnd.Next(races.Length);
            npc.Race = races[rIndex];
            if(gIndex == 0) {
                int nIndex = rnd.Next(maleNames.Length);
                npc.Name = maleNames[nIndex];
            } else {
                int nIndex = rnd.Next(femaleNames.Length);
                npc.Name = femaleNames[nIndex];
            }
            int sIndex = rnd.Next(surnames.Length);
            npc.Surname = surnames[sIndex];
            int cIndex = rnd.Next(classes.Length);
            npc.Class = classes[cIndex];
            int pIndex = rnd.Next(professions.Length);
            npc.Profession = professions[pIndex];
            int tIndex = rnd.Next(traits.Length);
            npc.Trait = traits[tIndex];
            npcs.Add(npc);
            Console.WriteLine("Testing this NPC: " + npc.Name + " " + npc.Surname + " " + npc.Gender + " " + npc.Class + " " + npc.Profession + " " + npc.Trait + " " + npc.Race);
        }
    }
}
