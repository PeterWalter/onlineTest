// See https://aka.ms/new-console-template for more information
using Cetap_Classes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2013.Word;
using Dumpify;
using System.Collections.ObjectModel;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using static System.Runtime.InteropServices.JavaScript.JSType;


Console.WriteLine("Territorium File reader");
Console.WriteLine("=============================================");
Console.WriteLine();

Console.Write("Enter file type to process :");
string filetype = Console.ReadLine();
filetype = filetype.ToUpper();
Console.Write("Enter name of File to process including the directory path");
string filepath = Console.ReadLine();
//string filepath = "D:/test/CEA_MAT_May_11.csv";
Console.WriteLine(".......................................");
Console.WriteLine(filepath);
// Create a helper function to handle repeated selection and filtering for sections
//private List<T> GetSectionData<T>(Collection<Read_online_AQL> my_recs, string sectionName, Func<Read_online_AQL, T> selector)
//{
//    return my_recs
//        .Where(t => t.Test.Contains(sectionName) || t.Test.Contains($"Afdeling {sectionName.Last()}"))
//        .Select(selector)
//        .ToList();
//}

try
{
    Loadrecs reader = new Loadrecs(filepath,filetype);
    switch (filetype)
    {
        case "AQL":
            Collection<Read_online_AQL> my_recs = reader.AQLrecs;

            // Define required sections
            List<string> requiredSections = new List<string>
                {
                    "Section 1", "Section 2", "Section 3", "Section 4", "Section 5", "Section 6", "Section 7"
                };

            // Process records to ensure all sections exist
            Collection<Read_online_AQL> processedRecords = new Collection<Read_online_AQL>();

            // Group records by UID
            var groupedWriters = my_recs.GroupBy(r => r.UID);

            foreach (var group in groupedWriters)
            {
                // Get existing sections for this UID
                var existingSections = group.Select(r => r.Test).ToList();
              
                // Add existing records to the processed list
                foreach (var record in group)
                {
                    processedRecords.Add(record);
                  //  my_test =record.Test;
                }
                // Identify and add missing sections
                foreach (string section in requiredSections)
                {
                    if (!existingSections.Any(s => s.Contains(section)))
                    {
                        string[] parts = group.First().Test.Split(':');
                        // Create a default record for the missing section
                        processedRecords.Add(new Read_online_AQL
                        {
                            Period = group.First().Period,
                            UID = group.Key, // Set UID for the group
                            Group = group.First().Group,
                            StartTest = group.First().StartTest,
                            Test = $"{parts[0]}: {section}", // Section name
                            RefNo = group.First().RefNo, // Default RefNo
                            Name = group.First().Name ?? "Unknown",
                            Surname = group.First().Surname ?? "Unknown",
                            Email = group.First().Email ?? "Unknown",
                            DOT = group.First().DOT ?? DateTime.Now.ToString(),
                            Question1 = 'X',
                            Question2 = 'X',
                            Question3 = 'X',
                            Question4 = 'X',
                            Question5 = 'X',
                            Question6 = 'X',
                            Question7 = 'X',
                            Question8 = 'X',
                            Question9 = 'X',
                            Question10 = 'X',
                            Question11 = 'X',
                            Question12 = 'X',
                            Question13 = 'X',
                            Question14 = 'X',
                            Question15 = 'X',
                            Question16 = 'X',
                            Question17 = 'X',
                            Question18 = 'X',
                            Question19 = 'X',
                            Question20 = 'X',
                            Question21 = 'X',
                            Question22 = 'X',
                            Question23 = 'X',
                            Question24 = 'X',
                            Question25 = 'X'
                        });
                    }

                }
                
            }
// ------------------------------------------------------------------------------------------------------------------
//--------------------------------------------------------------------------------------------------------------------
// new way of creating onlineAQL
//--------------------------------------------------------------------------------------------------------------------
       //    my_recs.Clear();
            var writers = processedRecords.GroupBy(r => r.UID);
            // add all processed records to the my_recs collection
            Collection<onlineAQL> processedData = new Collection<onlineAQL>();

            foreach (var Writer in writers )
            {
                var scoreRec = new onlineAQL();
                scoreRec.RefNo = Convert.ToInt64(Writer.First().RefNo);
                scoreRec.Name = Writer.First().Name;
                scoreRec.Surname = Writer.First().Surname;
                scoreRec.DOT = Writer.First().DOT;
                scoreRec.Group = Writer.First().Group;

                
                foreach (var record in Writer)
                {
                    string aSection = record.Test;
                    string[] parts = aSection.Split(':');

                    switch (parts[1].Trim())
                    {
                        case "Section 1":
                            scoreRec.Section1 = new char[17];
                            scoreRec.Section1[0] = record.Question1;
                            scoreRec.Section1[1] = record.Question2;
                            scoreRec.Section1[2] = record.Question3;    
                            scoreRec.Section1[3] = record.Question4;    
                            scoreRec.Section1[4] = record.Question5;
                            scoreRec.Section1[5] = record.Question6;
                            scoreRec.Section1[6] = record.Question7;
                            scoreRec.Section1[7] = record.Question8;
                            scoreRec.Section1[8] = record.Question9;
                            scoreRec.Section1[9] = record.Question10;
                            scoreRec.Section1[10] = record.Question11;
                            scoreRec.Section1[11] = record.Question12;
                            scoreRec.Section1[12] = record.Question13;
                            scoreRec.Section1[13] = record.Question14;
                            scoreRec.Section1[14] = record.Question15;
                            scoreRec.Section1[15] = record.Question16;
                            scoreRec.Section1[16] = record.Question17;
                            break;
                        case "Section 2":
                            scoreRec.Section2 = new char[17];
                            scoreRec.Section2[0] = record.Question1;
                            scoreRec.Section2[1] = record.Question2;
                            scoreRec.Section2[2] = record.Question3;
                            scoreRec.Section2[3] = record.Question4;
                            scoreRec.Section2[4] = record.Question5;
                            scoreRec.Section2[5] = record.Question6;
                            scoreRec.Section2[6] = record.Question7;
                            scoreRec.Section2[7] = record.Question8;
                            scoreRec.Section2[8] = record.Question9;
                            scoreRec.Section2[9] = record.Question10;
                            scoreRec.Section2[10] = record.Question11;
                            scoreRec.Section2[11] = record.Question12;
                            scoreRec.Section2[12] = record.Question13;
                            scoreRec.Section2[13] = record.Question14;
                            scoreRec.Section2[14] = record.Question15;
                            scoreRec.Section2[15] = record.Question16;
                            scoreRec.Section2[16] = record.Question17;
                            break;
                        case "Section 3":
                            scoreRec.Section3 = new char[25];
                            scoreRec.Section3[0] = record.Question1;
                            scoreRec.Section3[1] = record.Question2;
                            scoreRec.Section3[2] = record.Question3;
                            scoreRec.Section3[3] = record.Question4;
                            scoreRec.Section3[4] = record.Question5;
                            scoreRec.Section3[5] = record.Question6;
                            scoreRec.Section3[6] = record.Question7;
                            scoreRec.Section3[7] = record.Question8;
                            scoreRec.Section3[8] = record.Question9;
                            scoreRec.Section3[9] = record.Question10;
                            scoreRec.Section3[10] = record.Question11;
                            scoreRec.Section3[11] = record.Question12;
                            scoreRec.Section3[12] = record.Question13;
                            scoreRec.Section3[13] = record.Question14;
                            scoreRec.Section3[14] = record.Question15;
                            scoreRec.Section3[15] = record.Question16;
                            scoreRec.Section3[16] = record.Question17;
                            scoreRec.Section3[17] = record.Question18;
                            scoreRec.Section3[18] = record.Question19;
                            scoreRec.Section3[19] = record.Question20;
                            scoreRec.Section3[20] = record.Question21;
                            scoreRec.Section3[21] = record.Question22;
                            scoreRec.Section3[22] = record.Question23;
                            scoreRec.Section3[23] = record.Question24;
                            scoreRec.Section3[24] = record.Question25;
                            break;
                        case "Section 4":
                            scoreRec.Section4 = new char[25];
                            scoreRec.Section4[0] = record.Question1;
                            scoreRec.Section4[1] = record.Question2;
                            scoreRec.Section4[2] = record.Question3;
                            scoreRec.Section4[3] = record.Question4;
                            scoreRec.Section4[4] = record.Question5;
                            scoreRec.Section4[5] = record.Question6;
                            scoreRec.Section4[6] = record.Question7;
                            scoreRec.Section4[7] = record.Question8;
                            scoreRec.Section4[8] = record.Question9;
                            scoreRec.Section4[9] = record.Question10;
                            scoreRec.Section4[10] = record.Question11;
                            scoreRec.Section4[11] = record.Question12;
                            scoreRec.Section4[12] = record.Question13;
                            scoreRec.Section4[13] = record.Question14;
                            scoreRec.Section4[14] = record.Question15;
                            scoreRec.Section4[15] = record.Question16;
                            scoreRec.Section4[16] = record.Question17;
                            scoreRec.Section4[17] = record.Question18;
                            scoreRec.Section4[18] = record.Question19;
                            scoreRec.Section4[19] = record.Question20;
                            scoreRec.Section4[20] = record.Question21;
                            scoreRec.Section4[21] = record.Question22;
                            scoreRec.Section4[22] = record.Question23;
                            scoreRec.Section4[23] = record.Question24;
                            scoreRec.Section4[24] = record.Question25;

                            break;
                        case "Section 5":
                            scoreRec.Section5 = new char[17];
                            scoreRec.Section5[0] = record.Question1;
                            scoreRec.Section5[1] = record.Question2;
                            scoreRec.Section5[2] = record.Question3;
                            scoreRec.Section5[3] = record.Question4;
                            scoreRec.Section5[4] = record.Question5;
                            scoreRec.Section5[5] = record.Question6;
                            scoreRec.Section5[6] = record.Question7;
                            scoreRec.Section5[7] = record.Question8;
                            scoreRec.Section5[8] = record.Question9;
                            scoreRec.Section5[9] = record.Question10;
                            scoreRec.Section5[10] = record.Question11;
                            scoreRec.Section5[11] = record.Question12;
                            scoreRec.Section5[12] = record.Question13;
                            scoreRec.Section5[13] = record.Question14;
                            scoreRec.Section5[14] = record.Question15;
                            scoreRec.Section5[15] = record.Question16;
                            scoreRec.Section5[16] = record.Question17;
                            break;
                        case "Section 6":
                            scoreRec.Section6 = new char[24];
                            scoreRec.Section6[0] = record.Question1;
                            scoreRec.Section6[1] = record.Question2;
                            scoreRec.Section6[2] = record.Question3;
                            scoreRec.Section6[3] = record.Question4;
                            scoreRec.Section6[4] = record.Question5;
                            scoreRec.Section6[5] = record.Question6;
                            scoreRec.Section6[6] = record.Question7;
                            scoreRec.Section6[7] = record.Question8;
                            scoreRec.Section6[8] = record.Question9;
                            scoreRec.Section6[9] = record.Question10;
                            scoreRec.Section6[10] = record.Question11;
                            scoreRec.Section6[11] = record.Question12;
                            scoreRec.Section6[12] = record.Question13;
                            scoreRec.Section6[13] = record.Question14;
                            scoreRec.Section6[14] = record.Question15;
                            scoreRec.Section6[15] = record.Question16;
                            scoreRec.Section6[16] = record.Question17;
                            scoreRec.Section6[17] = record.Question18;
                            scoreRec.Section6[18] = record.Question19;
                            scoreRec.Section6[19] = record.Question20;
                            scoreRec.Section6[20] = record.Question21;
                            scoreRec.Section6[21] = record.Question22;
                            scoreRec.Section6[22] = record.Question23;
                            scoreRec.Section6[23] = record.Question24;
                           // scoreRec.Section6[24] = record.Question25;

                            break;
                        case "Section 7":
                            scoreRec.Section7 = new char[25];
                            scoreRec.Section7[0] = record.Question1;
                            scoreRec.Section7[1] = record.Question2;
                            scoreRec.Section7[2] = record.Question3;
                            scoreRec.Section7[3] = record.Question4;
                            scoreRec.Section7[4] = record.Question5;
                            scoreRec.Section7[5] = record.Question6;
                            scoreRec.Section7[6] = record.Question7;
                            scoreRec.Section7[7] = record.Question8;
                            scoreRec.Section7[8] = record.Question9;
                            scoreRec.Section7[9] = record.Question10;
                            scoreRec.Section7[10] = record.Question11;
                            scoreRec.Section7[11] = record.Question12;
                            scoreRec.Section7[12] = record.Question13;
                            scoreRec.Section7[13] = record.Question14;
                            scoreRec.Section7[14] = record.Question15;
                            scoreRec.Section7[15] = record.Question16;
                            scoreRec.Section7[16] = record.Question17;
                            scoreRec.Section7[17] = record.Question18;
                            scoreRec.Section7[18] = record.Question19;
                            scoreRec.Section7[19] = record.Question20;
                            scoreRec.Section7[20] = record.Question21;
                            scoreRec.Section7[21] = record.Question22;
                            scoreRec.Section7[22] = record.Question23;
                            scoreRec.Section7[23] = record.Question24;
                            scoreRec.Section7[24] = record.Question25;

                            break;

                    }
                }
              
                scoreRec.changeData();
                scoreRec.RecalculateCounts();
                processedData.Add(scoreRec);
            }


            //            // Use the helper function to fetch section data
            //            // change my_recs to processedRecords

            //              var section1 = GetSectionData(processedRecords, "Section 1", t => new
            //                    {
            //                        t.RefNo,
            //                        t.Surname,
            //                        t.Name,
            //                        t.Group,
            //                        t.Period,
            //                        t.DOT,
            //                        t.Email,
            //                        t.StartTest,
            //                        Questions = new[] { t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9, t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15, t.Question16, t.Question17 }
            //                    });

            //            var section2 = GetSectionData(processedRecords, "Section 2", t => new
            //            {
            //                t.RefNo,
            //                Questions = new[] { t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9, t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15, t.Question16, t.Question17 }
            //            });

            //            var section3 = GetSectionData(processedRecords, "Section 3", t => new
            //            {
            //                t.RefNo,
            //                Questions = new[] { t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9, t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15, t.Question16, t.Question17, t.Question18, t.Question19, t.Question20, t.Question21, t.Question22, t.Question23, t.Question24, t.Question25 }
            //            });
            //            // Repeat for other sections (section4, section5, etc.)
            //            var section4 = GetSectionData(processedRecords, "Section 4", t => new
            //            {
            //                t.RefNo,
            //                Questions = new[] { t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9, t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15, t.Question16, t.Question17, t.Question18, t.Question19, t.Question20, t.Question21, t.Question22, t.Question23, t.Question24, t.Question25 }
            //            });

            //            var section5 = GetSectionData(processedRecords, "Section 5", t => new
            //            {
            //                t.RefNo,
            //                Questions = new[] { t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9, t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15, t.Question16, t.Question17 }
            //            });

            //            var section6 = GetSectionData(processedRecords, "Section 6", t => new
            //            {
            //                t.RefNo,
            //                Questions = new[] { t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9, t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15, t.Question16, t.Question17, t.Question18, t.Question19, t.Question20, t.Question21, t.Question22, t.Question23, t.Question24 }
            //            });

            //            var section7 = GetSectionData(processedRecords, "Section 7", t => new
            //            {
            //                t.RefNo,
            //                Questions = new[] { t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9, t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15, t.Question16, t.Question17, t.Question18, t.Question19, t.Question20, t.Question21, t.Question22, t.Question23, t.Question24, t.Question25 }
            //            });


            //            // Join sections into one record using unique RefNo
            //            var AQL = section1
            //                .Join(section2, s1 => s1.RefNo, s2 => s2.RefNo, (s1, s2) => new { s1, s2 })
            //                .Join(section3, s12 => s12.s1.RefNo, s3 => s3.RefNo, (s12, s3) => new { s12, s3 })
            //                .Join(section4, s123 => s123.s12.s1.RefNo, s4 => s4.RefNo, (s123, s4) => new { s123, s4 })
            //                .Join(section5, s1234 => s1234.s123.s12.s1.RefNo, s5 => s5.RefNo, (s1234, s5) => new { s1234, s5 })
            //                .Join(section6, s12345 => s12345.s1234.s123.s12.s1.RefNo, s6 => s6.RefNo, (s12345, s6) => new { s12345, s6 })
            //                .Join(section7, s123456 => s123456.s12345.s1234.s123.s12.s1.RefNo, s7 => s7.RefNo, (s123456, s7) => new { s123456, s7 })
            //                .Select(t => new
            //                {
            //                    Person = t.s123456.s12345.s1234.s123.s12.s1,
            //                    sec2 = t.s123456.s12345.s1234.s123.s12.s2,
            //                    sec3 = t.s123456.s12345.s1234.s123.s3,
            //                    sec4 = t.s123456.s12345.s1234.s4,
            //                    sec5 = t.s123456.s12345.s5,
            //                    sec6 = t.s123456.s6,
            //                    sec7 = t.s7
            //                }).Distinct()
            //                .ToList();

            //            int count = 0;
            //            long NBT = 0;
            //            List<onlineAQL> AQLs = new List<onlineAQL>();

            //            foreach (var line in AQL)
            //            {
            //                var my_aql = new onlineAQL
            //                {
            //                    RefNo = Convert.ToInt64(line.Person.RefNo),
            //                    Name = line.Person.Name,
            //                    Surname = line.Person.Surname,
            //                    DOT = line.Person.DOT,
            //                    Group = line.Person.Group,
            //                };

            //                for (int i = 0; i < 25; i++)
            //                {

            //                    if (i < 17)
            //                    {
            //                        my_aql.Section1[i] = line.Person.Questions[i];
            //                        my_aql.Section2[i] = line.sec2.Questions[i];
            //                    }

            //                    if (i < line.sec3.Questions.Length)
            //                    {
            //                        my_aql.Section3[i] = line.sec3.Questions[i];
            //                    }

            //                    // Handle other sections (section4, section5, etc.) similarly
            //                    if (i < line.sec4.Questions.Length) my_aql.Section4[i] = line.sec4.Questions[i];
            //                    if (i < line.sec5.Questions.Length) my_aql.Section5[i] = line.sec5.Questions[i];
            //                    if (i < line.sec6.Questions.Length) my_aql.Section6[i] = line.sec6.Questions[i];
            //                    if (i < line.sec7.Questions.Length) my_aql.Section7[i] = line.sec7.Questions[i];
            //                }

            //                my_aql.changeData();
            //                my_aql.RecalculateCounts();
            //                string my_data = my_aql.ToString();
            //                 Console.WriteLine(my_data);
            //                // Console.WriteLine($"chars = {my_data.Count()}");
            //                AQLs.Add(my_aql);
            //                count++;

            //            }

            ////-----------------------------------------------------------------//
            //// Helper function to handle section data retrieval
            //static List<T> GetSectionData<T>(Collection<Read_online_AQL> my_recs, string sectionName, Func<Read_online_AQL, T> selector)
            //            {
            //                return my_recs
            //                    .Where(t => t.Test.Contains(sectionName) || t.Test.Contains($"Afdeling {sectionName.Last()}"))
            //                    .Select(selector)
            //                    .ToList();
            //            }

            var scoredAQL_2025 = processedData.ToString();
    var scoredAQL = processedData
    .Select(n => new
    {
        section = n.RefNo,
        sections = new Dictionary<string, int>()
        {
            { "sect1", n.Section1.Length == n.Section1Count ? n.Section1Count : -1 },
            { "sect2", n.Section2.Length == n.Section2Count ? n.Section2Count : -1 },
            { "section3", n.Section3.Length == n.Section3Count ? n.Section3Count : -1 },
            { "sect4", n.Section4.Length == n.Section4Count ? n.Section4Count : -1 },
            { "sect5", n.Section5.Length == n.Section5Count ? n.Section5Count : -1 },
            { "sect6", n.Section6.Length == n.Section6Count ? n.Section6Count : -1 },
            { "sect7", n.Section7.Length == n.Section7Count ? n.Section7Count : -1 }
        }
        .Where(s => s.Value != -1) // Exclude invalid (non-matching) sections
        .ToDictionary(s => s.Key, s => s.Value)
    })
    .Where(n => n.sections.Any()) // Include only if at least one section is valid
    .ToList();


     var removeAQL = processedData
    .Select(n => new
    {
        section = n.RefNo,
        sect1 = n.Section1.Length == n.Section1Count ? n.Section1Count : (int?)null,
        sect2 = n.Section2.Length == n.Section2Count ? n.Section2Count : (int?)null,
        sect3 = n.Section3.Length == n.Section3Count ? n.Section3Count : (int?)null,
        sect4 = n.Section4.Length == n.Section4Count ? n.Section4Count : (int?)null,
        sect5 = n.Section5.Length == n.Section5Count ? n.Section5Count : (int?)null,
        sect6 = n.Section6.Length == n.Section6Count ? n.Section6Count : (int?)null,
        sect7 = n.Section7.Length == n.Section7Count ? n.Section7Count : (int?)null
    })
    .Where(n =>
        // Condition 1: Both sect1 and sect2 must have values
        (n.sect1.HasValue && n.sect2.HasValue && n.sect5.HasValue && n.sect6.HasValue) ||
        // Condition 2: Both sect3 and sect4 must have values
        (n.sect3.HasValue && n.sect4.HasValue)
    )
    .ToList();

   // Convert the anonymous type list to a list of strings
   var removeAQLStrings = removeAQL
                .Select(n => $"{n.section}, {n.sect1}, {n.sect2}, {n.sect3}, {n.sect4}, {n.sect5}, {n.sect6}, {n.sect7}")
                .ToList();

    string nbt_Invalid = Path.Combine(Path.GetDirectoryName(filepath), "NBT_Invalid.txt");
    File.WriteAllLines(nbt_Invalid, removeAQLStrings);

            var groupedData = processedData
                                .GroupBy(t  => new
                                {
                                    AdjustedGroup = t.Group.Split(' ').Length == 3 ? string.Join(" ",t.Group.Split(' ').SkipLast(1)) : t.Group
                                })
                                .Select( g => new
                                {
                                    Group = g.Key.AdjustedGroup,
                                    Details = g.Select(x => x.ToString()).ToList()

                                }). ToList();

            foreach ( var group in groupedData )
            {
                string outfilename = $"{group.Group}.txt";
                string directory = Path.GetDirectoryName(filepath);
                string AQLfile = Path.Combine(directory, outfilename);
                File.WriteAllLines(AQLfile, group.Details);
            }
            
            var nbts = processedData
                      .Select( n => n.RefNo.ToString()).ToList();

            string nbtFile = Path.Combine(Path.GetDirectoryName(filepath), "NBT_References.txt");
            File.WriteAllLines(nbtFile, nbts);


            break;
        case "MAT":
            Collection<Read_online_MAT> mat_recs;
            List<onlineMAT> MATs = new List<onlineMAT>();
            mat_recs = new Collection<Read_online_MAT>();
            mat_recs = reader.MATrecs;

            foreach (var record in mat_recs)
            {
                onlineMAT my_mat = new onlineMAT();
                my_mat.RefNo = Convert.ToInt64(record.RefNo);
                my_mat.Name = record.Name;
                my_mat.Surname = record.Surname;
                my_mat.DOT = record.DOT;
                my_mat.Group = record.Group;
                for (int i = 0; i < 60; i++)
                {
                    switch (i)
                    {
                        case 0:
                            my_mat.Section[i] = record.Question1;
                            continue;
                        case 1:
                            my_mat.Section[i] = record.Question2;
                            continue;
                        case 2:
                            my_mat.Section[i] = record.Question3;
                            continue;
                        case 3:
                            my_mat.Section[i] = record.Question4;
                            continue;
                        case 4:
                            my_mat.Section[i] = record.Question5;
                            continue;
                        case 5:
                            my_mat.Section[i] = record.Question6;
                            continue;
                        case 6:
                            my_mat.Section[i] = record.Question7;
                            continue;
                        case 7:
                            my_mat.Section[i] = record.Question8;
                            continue;
                        case 8:
                            my_mat.Section[i] = record.Question9;
                            continue;
                        case 9:
                            my_mat.Section[i] = record.Question10;
                            continue;
                        case 10:
                            my_mat.Section[i] = record.Question11;
                            continue;
                        case 11:
                            my_mat.Section[i] = record.Question12;
                            continue;
                        case 12:
                            my_mat.Section[i] = record.Question13;
                            continue;
                        case 13:
                            my_mat.Section[i] = record.Question14;
                            continue;
                        case 14:
                            my_mat.Section[i] = record.Question15;
                            continue;
                        case 15:
                            my_mat.Section[i] = record.Question16;
                            continue;
                        case 16:
                            my_mat.Section[i] = record.Question17;
                            continue;
                        case 17:
                            my_mat.Section[i] = record.Question18;
                            continue;
                        case 18:
                            my_mat.Section[i] = record.Question19;
                            continue;
                        case 19:
                            my_mat.Section[i] = record.Question20;
                            continue;
                        case 20:
                            my_mat.Section[i] = record.Question21;
                            continue;
                        case 21:
                            my_mat.Section[i] = record.Question22;
                            continue;
                        case 22:
                            my_mat.Section[i] = record.Question23;
                            continue;
                        case 23:
                            my_mat.Section[i] = record.Question24;
                            continue;
                        case 24:
                            my_mat.Section[i] = record.Question25;
                            continue;
                        case 25:
                            my_mat.Section[i] = record.Question26;
                            continue;
                        case 26:
                            my_mat.Section[i] = record.Question27;
                            continue;
                        case 27:
                            my_mat.Section[i] = record.Question28;
                            continue;
                        case 28:
                            my_mat.Section[i] = record.Question29;
                            continue;
                        case 29:
                            my_mat.Section[i] = record.Question30;
                            continue;
                        case 30:
                            my_mat.Section[i] = record.Question31;
                            continue;
                        case 31:
                            my_mat.Section[i] = record.Question32;
                            continue;
                        case 32:
                            my_mat.Section[i] = record.Question33;
                            continue;
                        case 33:
                            my_mat.Section[i] = record.Question34;
                            continue;
                        case 34:
                            my_mat.Section[i] = record.Question35;
                            continue;
                        case 35:
                            my_mat.Section[i] = record.Question36;
                            continue;
                        case 36:
                            my_mat.Section[i] = record.Question37;
                            continue;
                        case 37:
                            my_mat.Section[i] = record.Question38;
                            continue;
                        case 38:
                            my_mat.Section[i] = record.Question39;
                            continue;
                        case 39:
                            my_mat.Section[i] = record.Question40;
                            continue;
                        case 40:
                            my_mat.Section[i] = record.Question41;
                            continue;
                        case 41:
                            my_mat.Section[i] = record.Question42;
                            continue;
                        case 42:
                            my_mat.Section[i] = record.Question43;
                            continue;
                        case 43:
                            my_mat.Section[i] = record.Question44;
                            continue;
                        case 44:
                            my_mat.Section[i] = record.Question45;
                            continue;
                        case 45:
                            my_mat.Section[i] = record.Question46;
                            continue;
                        case 46:
                            my_mat.Section[i] = record.Question47;
                            continue;
                        case 47:
                            my_mat.Section[i] = record.Question48;
                            continue;
                        case 48:
                            my_mat.Section[i] = record.Question49;
                            continue;
                        case 49:
                            my_mat.Section[i] = record.Question50;
                            continue;
                        case 50:
                            my_mat.Section[i] = record.Question51;
                            continue;
                        case 51:
                            my_mat.Section[i] = record.Question52;
                            continue;
                        case 52:
                            my_mat.Section[i] = record.Question53;
                            continue;
                        case 53:
                            my_mat.Section[i] = record.Question54;
                            continue;
                        case 54:
                            my_mat.Section[i] = record.Question55;
                            continue;
                        case 55:
                            my_mat.Section[i] = record.Question56;
                            continue;
                        case 56:
                            my_mat.Section[i] = record.Question57;
                            continue;
                        case 57:
                            my_mat.Section[i] = record.Question58;
                            continue;
                        case 58:
                            my_mat.Section[i] = record.Question59;
                            continue;
                        case 59:
                            my_mat.Section[i] = record.Question60;
                            break;

                        default:
                            break;

                    }
                }
                my_mat.changeData();
                MATs.Add(my_mat);

            }

            var myData = MATs
                                .GroupBy(t => new
                                {
                                    AdjustedGroup = t.Group.Split(' ').Length == 2 ? string.Join(" ", t.Group.Split(' ').SkipLast(1)) : t.Group
                                })
                                .Select(g => new
                                {
                                    Group = g.Key.AdjustedGroup,
                                    Details = g.Select(x => x.ToString()).ToList()

                                }).ToList();

            foreach (var group in myData)
            {
                string outfilename = $"{group.Group}.txt";
                string directory = Path.GetDirectoryName(filepath);
                string MATfile = Path.Combine(directory, outfilename);
                File.WriteAllLines(MATfile, group.Details);
            }

            var MATA = MATs
            .Where(t => t.Group.Contains("MA"))
            .ToList();

            string MATA_test = MATA.FirstOrDefault()?.Group;
            string MATA_part = MATA_test.Trim().Substring(0, 2);
            int MAtestcode = Convert.ToInt32(MATA_test.Trim().Substring(2));
            string MATAfileName = "MATA" + MAtestcode.ToString("D4") + MATA_part + MAtestcode.ToString() + ".txt";

            var MATE = MATs
            .Where(t => t.Group.Contains("MT"))
            .ToList();

            string MATE_test = MATE.FirstOrDefault()?.Group;
            string MATE_part = MATE_test.Trim().Substring(0, 2);
            int MEtestcode = Convert.ToInt32(MATE_test.Trim().Substring(2));
            string MATEfileName = "MATE" + MEtestcode.ToString("D4") + MATE_part + MEtestcode.ToString() + ".txt";

            // Writing records to file
            string directory1 = Path.GetDirectoryName(filepath);
            string MATEfile = Path.Combine(directory1, MATEfileName);
            string MATAfile = Path.Combine(directory1, MATAfileName);
            using (StreamWriter writer = new StreamWriter(MATEfile))
            {
                foreach (var record in MATE)
                {
                    writer.WriteLine(record.ToString());
                }
            }
            using (StreamWriter writer = new StreamWriter(MATAfile))
            {
                foreach (var record in MATA)
                {
                    writer.WriteLine(record.ToString());
                }
            }

            break;
        // BIO
        case "BIO":
            Collection<Read_online_BIO> bio_recs;
            List<onlineBIO> BIOs = new List<onlineBIO>();
            bio_recs = new Collection<Read_online_BIO>();
            bio_recs = reader.BIOrecs;

            var citizen = bio_recs
            .Where(t => t.Question.Contains("CITIZENSHIP"))
            .ToList();
            var classification = bio_recs
            .Where(t => t.Question.Contains("CLASSIFICATION"))
            .ToList();

            var faculty1 = bio_recs
            .Where(t => t.Question.Contains("FACULTY to which you have applied first"))
            .ToList();
            var faculty2 = bio_recs
            .Where(t => t.Question.Contains("FACULTY to which you have applied second"))
            .ToList();
            var faculty3 = bio_recs
            .Where(t => t.Question.Contains("FACULTY to which you have applied third"))
            .ToList();
            var gender = bio_recs
            .Where(t => t.Question.Contains("GENDER"))
            .ToList();
            var hlanguage = bio_recs
            .Where(t => t.Question.Contains("HOME LANGUAGE"))
            .ToList();
            var idtype = bio_recs
            .Where(t => t.Question.Contains("ID TYPE"))
            .ToList();
            var slanguage = bio_recs
            .Where(t => t.Question.Contains("SCHOOL LANGUAGE"))
            .ToList();

            var person_bio = citizen
                        .Join(classification, c => c.RefNo, c2 => c2.RefNo, (c, c2) => new { Citizen = c, classify = c2.Answer })
                        .Join(faculty1, c => c.Citizen.RefNo, c3 => c3.RefNo, (c, c3) => new { Citizen = c, faculty1 = c3.Answer })
                        .Join(faculty2, c => c.Citizen.Citizen.RefNo, c4 => c4.RefNo, (c, c4) => new { Citizen = c, faculty2 = c4.Answer })
                        .Join(faculty3, c => c.Citizen.Citizen.Citizen.RefNo, c5 => c5.RefNo, (c, c5) => new { Citizen = c, faculty3 = c5.Answer })
                        .Join(gender, c => c.Citizen.Citizen.Citizen.Citizen.RefNo, c6 => c6.RefNo, (c, c6) => new { Citizen = c, gender = c6.Answer })
                        .Join(hlanguage, c => c.Citizen.Citizen.Citizen.Citizen.Citizen.RefNo, c7 => c7.RefNo, (c, c7) => new { Citizen = c, hlang = c7.Answer })
                        .Join(idtype, c => c.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.RefNo, c8 => c8.RefNo, (c, c8) => new { Citizen = c, idtype = c8.Answer })
                        .Join(slanguage, c => c.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.RefNo, c9 => c9.RefNo, (c, c9) => new { Citizen = c, slang = c9.Answer })
                        .Select(t => new
                        {
                            t.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen,
                            t.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.classify,
                            t.Citizen.Citizen.Citizen.Citizen.Citizen.Citizen.faculty1,
                            t.Citizen.Citizen.Citizen.Citizen.Citizen.faculty2,
                            t.Citizen.Citizen.Citizen.Citizen.faculty3,
                            t.Citizen.Citizen.Citizen.gender,
                            t.Citizen.Citizen.hlang,
                            t.Citizen.idtype,
                            t.slang
                        }).ToList();

            var query = from c in citizen
                             join c1 in classification on c.RefNo equals c1.RefNo
                             join f1 in faculty1 on c.RefNo equals f1.RefNo
                             join f2 in faculty2 on c.RefNo equals f2.RefNo
                             join f3 in faculty3 on c.RefNo equals f3.RefNo
                             join g in gender on c.RefNo equals g.RefNo
                             join h1 in hlanguage on c.RefNo equals h1.RefNo
                             join idt in idtype on c.RefNo equals idt.RefNo
                             join s1 in slanguage on c.RefNo equals s1.RefNo
                             select new
                             {
                                 Citizen = c,
                                 Classify = c1.Answer,
                                 faculty1 = f1.Answer,
                                 faculty2 = f2.Answer,
                                 faculty3 = f3.Answer,
                                 gender = g.Answer,
                                 hlang = h1.Answer,
                                 idtype = idt.Answer,
                                 slang = s1.Answer
                             };

            var p_bio = query.ToList();
            // reading every record in file
            foreach (var record in person_bio)
            {
                onlineBIO my_bio = new onlineBIO();
                my_bio.RefNo = Convert.ToInt64(record.Citizen.RefNo);
                my_bio.FullName = record.Citizen.Fullname;
                my_bio.Citizenship = record.Citizen.Answer == "South Africa" ? 1 : record.Citizen.Answer == "SADC country" ? 2 : record.Citizen.Answer == "Other African country" ? 3 : 4;
                my_bio.email = record.Citizen.Email;
                my_bio.Classification = record.classify == "Black" ? 1 : record.classify == "Coloured" ? 2 : record.classify == "Indian/Asian" ? 3 : record.classify == "White" ? 4 : 5;

                my_bio.Faculty1 = record.faculty1 == "Allied Healthcare / Nursing" ? 'A' : record.faculty1 == "Art / Design" ? 'B' : record.faculty1 == "Business / Commerce / Management" ? 'C' : record.faculty1 == "Education " ? 'D' : 
                    record.faculty1 == "Engineering / Built Environment" ? 'E' : record.faculty1 == "Health Sciences" ? 'Y' : record.faculty1 == "Hospitality / Tourism " ? 'G' : record.faculty1 == "Humanities " ? 'H' : 
                    record.faculty1 == "Information & Communication Technology " ? 'I' : record.faculty1 == "Law " ? 'J' : record.faculty1 == "Science / Mathematics " ? 'K' : 'L';
                
                my_bio.Faculty2 = record.faculty2 == "Allied Healthcare / Nursing" ? 'A' : record.faculty2 == "Art / Design" ? 'B' : record.faculty2 == "Business / Commerce / Management" ? 'C' : record.faculty2 == "Education " ? 'D' :
                    record.faculty2 == "Engineering / Built Environment" ? 'E' : record.faculty2 == "Health Sciences" ? 'Y' : record.faculty2 == "Hospitality / Tourism " ? 'G' : record.faculty2 == "Humanities " ? 'H' :
                    record.faculty2 == "Information & Communication Technology " ? 'I' : record.faculty2 == "Law " ? 'J' : record.faculty2 == "Science / Mathematics " ? 'K' : 'L';
                
                my_bio.Faculty3 = record.faculty3 == "Allied Healthcare / Nursing" ? 'A' : record.faculty3 == "Art / Design" ? 'B' : record.faculty3 == "Business / Commerce / Management" ? 'C' : record.faculty3 == "Education " ? 'D' :
                    record.faculty3 == "Engineering / Built Environment" ? 'E' : record.faculty3 == "Health Sciences" ? 'Y' : record.faculty3 == "Hospitality / Tourism " ? 'G' : record.faculty3 == "Humanities " ? 'H' :
                    record.faculty3 == "Information & Communication Technology " ? 'I' : record.faculty3 == "Law " ? 'J' : record.faculty3 == "Science / Mathematics " ? 'K' : 'L';
                
                my_bio.Gender = record.gender == "Male" ? 1 : 2;

                my_bio.HomeLanguage = record.hlang == "Afrikaans" ? 1 : record.hlang == "English" ? 2 : record.hlang == "isiNdebele" ? 3 : record.hlang == "isiXhosa" ? 4 : record.hlang == "isiZulu" ? 5 : record.hlang == "Sesotho" ? 6 :
                    record.hlang == "Sesotho sa Leboa" ? 7 : record.hlang == "Setswana" ? 8 : record.hlang == "siSwati" ? 9 : record.hlang == "Tshivenda" ? 10 : record.hlang == "Xitsonga" ? 11 : 12;
                
                my_bio.IdType = record.idtype == "South African" ? 1 : 2;
                my_bio.SchoolLanguage = record.slang == "Afrikaans" ? 1 : record.slang == "English" ? 2 : 3;
                my_bio.Finish_date = record.Citizen.FinishDate;
                my_bio.Group = record.Citizen.Group == "BIO Eng" ? 'E' : 'A';


            
                BIOs.Add(my_bio);
                                //
                //string my_data = my_bio.ToString();
                //Console.WriteLine(my_data);
            }
             string directory3 = Path.GetDirectoryName(filepath);
            string BIOfile = Path.Combine(directory3, "BIO.txt");
            using (StreamWriter writer = new StreamWriter(BIOfile))
            {
                foreach (var record in BIOs)
                {
                    writer.WriteLine(record.ToString());
                }
            }

            // generate the bio excel file for scoring
            Generate_Bio_Excel_file(BIOs);
            break;
        default:
            break;


    }

}
catch(FileNotFoundException)
{
    Console.WriteLine($"File '{filepath}' not found.");
}

void Generate_Bio_Excel_file(List<onlineBIO> writers)
{
    using (var workbook = new XLWorkbook())
    {
        var worksheet = workbook.Worksheets.Add("writers");
        var currentRow = 1;
        // Add headers
        worksheet.Cell(currentRow, 1).Value = "RefNo";
        worksheet.Cell(currentRow, 2).Value = "Barcode";
        worksheet.Cell(currentRow, 3).Value = "ID_NUMBER";
        worksheet.Cell(currentRow, 4).Value = "ID_Foreign";
        worksheet.Cell(currentRow, 5).Value = "ID_Type";
        worksheet.Cell(currentRow, 6).Value = "Gender";
        worksheet.Cell(currentRow, 7).Value = "Citizenship";
        worksheet.Cell(currentRow, 8).Value = "HEALTH_SCI_APP";
        worksheet.Cell(currentRow, 9).Value = "Date_of_Birth";
        worksheet.Cell(currentRow, 10).Value = "SURNAME";
        worksheet.Cell(currentRow, 11).Value = "FIRST_NAME";
        worksheet.Cell(currentRow, 12).Value = "INITALS";
        worksheet.Cell(currentRow, 13).Value = "Test_Cen_Code";
        worksheet.Cell(currentRow, 14).Value = "DATE";
        worksheet.Cell(currentRow, 15).Value = "Home_Lang";
        worksheet.Cell(currentRow, 16).Value = "GR12_Language";
        worksheet.Cell(currentRow, 17).Value = "Classification";
        worksheet.Cell(currentRow, 18).Value = "AQL_LANG";
        worksheet.Cell(currentRow, 19).Value = "AQL_CODE";
        worksheet.Cell(currentRow, 20).Value = "MAT_LANG";
        worksheet.Cell(currentRow, 21).Value = "MAT_CODE";
        worksheet.Cell(currentRow, 22).Value = "Faculty1";
        worksheet.Cell(currentRow, 23).Value = "Faculty2";
        worksheet.Cell(currentRow, 24).Value = "Faculty3";
        worksheet.Cell(currentRow, 25).Value = "AQL_TestNo";
        worksheet.Cell(currentRow, 26).Value = "MAT_TestNo";
        // Add data
        foreach (var person in writers)
        {
            currentRow++;
            worksheet.Cell(currentRow, 1).Value = person.RefNo;
            worksheet.Cell(currentRow, 5).Value = person.IdType;
            worksheet.Cell(currentRow, 6).Value = person.Gender;
            worksheet.Cell(currentRow, 7).Value = person.Citizenship;
            worksheet.Cell(currentRow, 8).Value = person.Faculty1.ToString();
            worksheet.Cell(currentRow, 10).Value = person.FullName;
            worksheet.Cell(currentRow, 13).Value = "99999";
            worksheet.Cell(currentRow, 15).Value = person.HomeLanguage;
            worksheet.Cell(currentRow, 16).Value = person.SchoolLanguage;
            worksheet.Cell(currentRow, 17).Value = person.Classification;
            worksheet.Cell(currentRow, 22).Value = person.Faculty1.ToString();
            worksheet.Cell(currentRow, 23).Value = person.Faculty2.ToString();
            worksheet.Cell(currentRow, 24).Value = person.Faculty3.ToString();
        }
        string directory5 = Path.GetDirectoryName(filepath);
        string excelfile = Path.Combine(directory5, "NBT_AnswershettBio.xlsx");
        workbook.SaveAs(excelfile);
    }
    Console.WriteLine("************************************************");
    Console.WriteLine("Bio Completed");
    Console.ReadLine();


}
