// See https://aka.ms/new-console-template for more information
using Cetap_Classes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Word;
using Dumpify;
using System.Collections.ObjectModel;


Console.WriteLine("Territorium File reader");
Console.WriteLine("=============================================");
Console.WriteLine();
Console.Write("Enter filename for output file:");
string filename = Console.ReadLine();
Console.Write("Enter file type to process :");
string filetype = Console.ReadLine();
filetype = filetype.ToUpper();
Console.Write("Enter name of File to process including the directory path");
string filepath = Console.ReadLine();
//string filepath = "D:/test/CEA_MAT_May_11.csv";
Console.WriteLine(".......................................");
Console.WriteLine(filepath);

try
{
    Loadrecs reader = new Loadrecs(filepath,filetype);
    switch (filetype)
    {
        case "AQL":
            Collection<Read_online_AQL> my_recs;
            List<onlineAQL> AQLs = new List<onlineAQL>();
            my_recs = new Collection<Read_online_AQL>();
            my_recs = reader.AQLrecs;
            var section1 = my_recs
        .Where(t => t.Test.Contains("Section 1") || t.Test.Contains("Afdeling 1"))
        .Select(t => new
        {
            t.RefNo,
            t.Test,
            t.Surname,
            t.Name,
            t.Group,
            t.Period,
            t.DOT,
            t.Email,
            t.StartTest,
            t.Question1,
            t.Question2,
            t.Question3,
            t.Question4,
            t.Question5,
            t.Question6,
            t.Question7,
            t.Question8,
            t.Question9,
            t.Question10,
            t.Question11,
            t.Question12,
            t.Question13,
            t.Question14,
            t.Question15,
            t.Question16,
            t.Question17
        })
        .ToList();

            var section2 = my_recs
           .Where(t => t.Test.Contains("Section 2") || t.Test.Contains("Afdeling 2"))
            .Select(t => new
            {
                t.RefNo,
                t.Question1,
                t.Question2,
                t.Question3,
                t.Question4,
                t.Question5,
                t.Question6,
                t.Question7,
                t.Question8,
                t.Question9,
                t.Question10,
                t.Question11,
                t.Question12,
                t.Question13,
                t.Question14,
                t.Question15,
                t.Question16,
                t.Question17
            })
           .ToList();

            var section3 = my_recs
            .Where(t => t.Test.Contains("Section 3") || t.Test.Contains("Afdeling 3"))
             .Select(t => new
             {
                 t.RefNo,
                 t.Question1,
                 t.Question2,
                 t.Question3,
                 t.Question4,
                 t.Question5,
                 t.Question6,
                 t.Question7,
                 t.Question8,
                 t.Question9,
                 t.Question10,
                 t.Question11,
                 t.Question12,
                 t.Question13,
                 t.Question14,
                 t.Question15,
                 t.Question16,
                 t.Question17,
                 t.Question18,
                 t.Question19,
                 t.Question20,
                 t.Question21,
                 t.Question22,
                 t.Question23,
                 t.Question24,
                 t.Question25
             })
            .ToList();

            var section4 = my_recs
            .Where(t => t.Test.Contains("Section 4") || t.Test.Contains("Afdeling 4"))
             .Select(t => new
             {
                 t.RefNo,
                 t.Question1,
                 t.Question2,
                 t.Question3,
                 t.Question4,
                 t.Question5,
                 t.Question6,
                 t.Question7,
                 t.Question8,
                 t.Question9,
                 t.Question10,
                 t.Question11,
                 t.Question12,
                 t.Question13,
                 t.Question14,
                 t.Question15,
                 t.Question16,
                 t.Question17,
                 t.Question18,
                 t.Question19,
                 t.Question20,
                 t.Question21,
                 t.Question22,
                 t.Question23,
                 t.Question24,
                 t.Question25
             })
            .ToList();

            var section5 = my_recs
            .Where(t => t.Test.Contains("Section 5") || t.Test.Contains("Afdeling 5"))
             .Select(t => new
             {
                 t.RefNo,
                 t.Question1,
                 t.Question2,
                 t.Question3,
                 t.Question4,
                 t.Question5,
                 t.Question6,
                 t.Question7,
                 t.Question8,
                 t.Question9,
                 t.Question10,
                 t.Question11,
                 t.Question12,
                 t.Question13,
                 t.Question14,
                 t.Question15,
                 t.Question16,
                 t.Question17
             })
            .ToList();

            var section6 = my_recs
            .Where(t => t.Test.Contains("Section 6") || t.Test.Contains("Afdeling 6"))
             .Select(t => new
             {
                 t.RefNo,
                 t.Question1,
                 t.Question2,
                 t.Question3,
                 t.Question4,
                 t.Question5,
                 t.Question6,
                 t.Question7,
                 t.Question8,
                 t.Question9,
                 t.Question10,
                 t.Question11,
                 t.Question12,
                 t.Question13,
                 t.Question14,
                 t.Question15,
                 t.Question16,
                 t.Question17,
                 t.Question18,
                 t.Question19,
                 t.Question20,
                 t.Question21,
                 t.Question22,
                 t.Question23,
                 t.Question24
             })
            .ToList();

            var section7 = my_recs
            .Where(t => t.Test.Contains("Section 7") || t.Test.Contains("Afdeling 7"))
            .Select(t => new
            {
                t.RefNo,
                t.Question1,
                t.Question2,
                t.Question3,
                t.Question4,
                t.Question5,
                t.Question6,
                t.Question7,
                t.Question8,
                t.Question9,
                t.Question10,
                t.Question11,
                t.Question12,
                t.Question13,
                t.Question14,
                t.Question15,
                t.Question16,
                t.Question17,
                t.Question18,
                t.Question19,
                t.Question20,
                t.Question21,
                t.Question22,
                t.Question23,
                t.Question24,
                t.Question25
            })
            .ToList();

            // join sections into one record using unique refNo
            var AQL = section1
                .Join(section2, s1 => s1.RefNo, s2 => s2.RefNo, (s1, s2) => new { Person = s1, sec2 = s2 })
                .Join(section3, s1 => s1.Person.RefNo, s3 => s3.RefNo, (s1, s3) => new { Person = s1, sec3 = s3 })

                .Join(section4, s1 => s1.Person.Person.RefNo, s4 => s4.RefNo, (s1, s4) => new { Person = s1, sec4 = s4 })

                .Join(section5, s1 => s1.Person.Person.Person.RefNo, s5 => s5.RefNo, (s1, s5) => new { Person = s1, sec5 = s5 })

                .Join(section6, s1 => s1.Person.Person.Person.Person.RefNo, s6 => s6.RefNo, (s1, s6) => new { Person = s1, sec6 = s6 })

                .Join(section7, s1 => s1.Person.Person.Person.Person.Person.RefNo, s7 => s7.RefNo, (s1, s7) => new { Person = s1, sec7 = s7 })

                .Select(t => new
                {
                    t.Person.Person.Person.Person.Person.Person,
                    t.Person.Person.Person.Person.Person.sec2,
                    t.Person.Person.Person.Person.sec3,
                    t.Person.Person.Person.sec4,
                    t.Person.Person.sec5,
                    t.Person.sec6,
                    t.sec7
                })
                .ToList();


            int count = 0;
            long NBT = 0;
            foreach (var line in AQL)
            {
                onlineAQL my_aql = new onlineAQL();
                my_aql.RefNo = Convert.ToInt64(line.Person.RefNo);
                my_aql.Name = line.Person.Name;
                my_aql.Surname = line.Person.Surname;
                my_aql.DOT = line.Person.DOT;
                my_aql.Group = line.Person.Group;
                for (int i = 0; i < 26; i++)
                {

                    switch (i)
                    {
                        case 0:
                            my_aql.Section1[i] = line.Person.Question1;
                            my_aql.Section2[i] = line.sec2.Question1;
                            my_aql.Section3[i] = line.sec3.Question1;
                            my_aql.Section4[i] = line.sec4.Question1;
                            my_aql.Section5[i] = line.sec5.Question1;
                            my_aql.Section6[i] = line.sec6.Question1;
                            my_aql.Section7[i] = line.sec7.Question1;
                            continue;
                        case 1:
                            my_aql.Section1[i] = line.Person.Question2;
                            my_aql.Section2[i] = line.sec2.Question2;
                            my_aql.Section3[i] = line.sec3.Question2;
                            my_aql.Section4[i] = line.sec4.Question2;
                            my_aql.Section5[i] = line.sec5.Question2;
                            my_aql.Section6[i] = line.sec6.Question2;
                            my_aql.Section7[i] = line.sec7.Question2;
                            continue;
                        case 2:
                            my_aql.Section1[i] = line.Person.Question3;
                            my_aql.Section2[i] = line.sec2.Question3;
                            my_aql.Section3[i] = line.sec3.Question3;
                            my_aql.Section4[i] = line.sec4.Question3;
                            my_aql.Section5[i] = line.sec5.Question3;
                            my_aql.Section6[i] = line.sec6.Question3;
                            my_aql.Section7[i] = line.sec7.Question3;
                            continue;
                        case 3:
                            my_aql.Section1[i] = line.Person.Question4;
                            my_aql.Section2[i] = line.sec2.Question4;
                            my_aql.Section3[i] = line.sec3.Question4;
                            my_aql.Section4[i] = line.sec4.Question4;
                            my_aql.Section5[i] = line.sec5.Question4;
                            my_aql.Section6[i] = line.sec6.Question4;
                            my_aql.Section7[i] = line.sec7.Question4;
                            continue;
                        case 4:
                            my_aql.Section1[i] = line.Person.Question5;
                            my_aql.Section2[i] = line.sec2.Question5;
                            my_aql.Section3[i] = line.sec3.Question5;
                            my_aql.Section4[i] = line.sec4.Question5;
                            my_aql.Section5[i] = line.sec5.Question5;
                            my_aql.Section6[i] = line.sec6.Question5;
                            my_aql.Section7[i] = line.sec7.Question5;
                            continue;
                        case 5:
                            my_aql.Section1[i] = line.Person.Question6;
                            my_aql.Section2[i] = line.sec2.Question6;
                            my_aql.Section3[i] = line.sec3.Question6;
                            my_aql.Section4[i] = line.sec4.Question6;
                            my_aql.Section5[i] = line.sec5.Question6;
                            my_aql.Section6[i] = line.sec6.Question6;
                            my_aql.Section7[i] = line.sec7.Question6;
                            continue;

                        case 6:
                            my_aql.Section1[i] = line.Person.Question7;
                            my_aql.Section2[i] = line.sec2.Question7;
                            my_aql.Section3[i] = line.sec3.Question7;
                            my_aql.Section4[i] = line.sec4.Question7;
                            my_aql.Section5[i] = line.sec5.Question7;
                            my_aql.Section6[i] = line.sec6.Question7;
                            my_aql.Section7[i] = line.sec7.Question7;
                            continue;
                        case 7:
                            my_aql.Section1[i] = line.Person.Question8;
                            my_aql.Section2[i] = line.sec2.Question8;
                            my_aql.Section3[i] = line.sec3.Question8;
                            my_aql.Section4[i] = line.sec4.Question8;
                            my_aql.Section5[i] = line.sec5.Question8;
                            my_aql.Section6[i] = line.sec6.Question8;
                            my_aql.Section7[i] = line.sec7.Question8;
                            continue;
                        case 8:
                            my_aql.Section1[i] = line.Person.Question9;
                            my_aql.Section2[i] = line.sec2.Question9;
                            my_aql.Section3[i] = line.sec3.Question9;
                            my_aql.Section4[i] = line.sec4.Question9;
                            my_aql.Section5[i] = line.sec5.Question9;
                            my_aql.Section6[i] = line.sec6.Question9;
                            my_aql.Section7[i] = line.sec7.Question9;
                            continue;
                        case 9:
                            my_aql.Section1[i] = line.Person.Question10;
                            my_aql.Section2[i] = line.sec2.Question10;
                            my_aql.Section3[i] = line.sec3.Question10;
                            my_aql.Section4[i] = line.sec4.Question10;
                            my_aql.Section5[i] = line.sec5.Question10;
                            my_aql.Section6[i] = line.sec6.Question10;
                            my_aql.Section7[i] = line.sec7.Question10;
                            continue;
                        case 10:
                            my_aql.Section1[i] = line.Person.Question11;
                            my_aql.Section2[i] = line.sec2.Question11;
                            my_aql.Section3[i] = line.sec3.Question11;
                            my_aql.Section4[i] = line.sec4.Question11;
                            my_aql.Section5[i] = line.sec5.Question11;
                            my_aql.Section6[i] = line.sec6.Question11;
                            my_aql.Section7[i] = line.sec7.Question11;
                            continue;
                        case 11:
                            my_aql.Section1[i] = line.Person.Question12;
                            my_aql.Section2[i] = line.sec2.Question12;
                            my_aql.Section3[i] = line.sec3.Question12;
                            my_aql.Section4[i] = line.sec4.Question12;
                            my_aql.Section5[i] = line.sec5.Question12;
                            my_aql.Section6[i] = line.sec6.Question12;
                            my_aql.Section7[i] = line.sec7.Question12;
                            continue;
                        case 12:
                            my_aql.Section1[i] = line.Person.Question13;
                            my_aql.Section2[i] = line.sec2.Question13;
                            my_aql.Section3[i] = line.sec3.Question13;
                            my_aql.Section4[i] = line.sec4.Question13;
                            my_aql.Section5[i] = line.sec5.Question13;
                            my_aql.Section6[i] = line.sec6.Question13;
                            my_aql.Section7[i] = line.sec7.Question13;
                            continue;
                        case 13:
                            my_aql.Section1[i] = line.Person.Question14;
                            my_aql.Section2[i] = line.sec2.Question14;
                            my_aql.Section3[i] = line.sec3.Question14;
                            my_aql.Section4[i] = line.sec4.Question14;
                            my_aql.Section5[i] = line.sec5.Question14;
                            my_aql.Section6[i] = line.sec6.Question14;
                            my_aql.Section7[i] = line.sec7.Question14;
                            continue;
                        case 14:
                            my_aql.Section1[i] = line.Person.Question15;
                            my_aql.Section2[i] = line.sec2.Question15;
                            my_aql.Section3[i] = line.sec3.Question15;
                            my_aql.Section4[i] = line.sec4.Question15;
                            my_aql.Section5[i] = line.sec5.Question15;
                            my_aql.Section6[i] = line.sec6.Question15;
                            my_aql.Section7[i] = line.sec7.Question15;
                            continue;
                        case 15:
                            my_aql.Section1[i] = line.Person.Question16;
                            my_aql.Section2[i] = line.sec2.Question16;
                            my_aql.Section3[i] = line.sec3.Question16;
                            my_aql.Section4[i] = line.sec4.Question16;
                            my_aql.Section5[i] = line.sec5.Question16;
                            my_aql.Section6[i] = line.sec6.Question16;
                            my_aql.Section7[i] = line.sec7.Question16;
                            continue;
                        case 16:
                            my_aql.Section1[i] = line.Person.Question17;
                            my_aql.Section2[i] = line.sec2.Question17;
                            my_aql.Section3[i] = line.sec3.Question17;
                            my_aql.Section4[i] = line.sec4.Question17;
                            my_aql.Section5[i] = line.sec5.Question17;
                            my_aql.Section6[i] = line.sec6.Question17;
                            my_aql.Section7[i] = line.sec7.Question17;
                            continue;
                        case 17:
                            my_aql.Section3[i] = line.sec3.Question18;
                            my_aql.Section4[i] = line.sec4.Question18;
                            my_aql.Section6[i] = line.sec6.Question18;
                            my_aql.Section7[i] = line.sec7.Question18;
                            continue;
                        case 18:
                            my_aql.Section3[i] = line.sec3.Question19;
                            my_aql.Section4[i] = line.sec4.Question19;
                            my_aql.Section6[i] = line.sec6.Question19;
                            my_aql.Section7[i] = line.sec7.Question19;
                            continue;
                        case 19:
                            my_aql.Section3[i] = line.sec3.Question20;
                            my_aql.Section4[i] = line.sec4.Question20;
                            my_aql.Section6[i] = line.sec6.Question20;
                            my_aql.Section7[i] = line.sec7.Question20;
                            continue;
                        case 20:
                            my_aql.Section3[i] = line.sec3.Question21;
                            my_aql.Section4[i] = line.sec4.Question21;
                            my_aql.Section6[i] = line.sec6.Question21;
                            my_aql.Section7[i] = line.sec7.Question21;
                            continue;
                        case 21:
                            my_aql.Section3[i] = line.sec3.Question22;
                            my_aql.Section4[i] = line.sec4.Question22;
                            my_aql.Section6[i] = line.sec6.Question22;
                            my_aql.Section6[i] = line.sec6.Question22;
                            my_aql.Section7[i] = line.sec7.Question22;
                            continue;
                        case 22:
                            my_aql.Section3[i] = line.sec3.Question23;
                            my_aql.Section4[i] = line.sec4.Question23;
                            my_aql.Section6[i] = line.sec6.Question23;
                            my_aql.Section7[i] = line.sec7.Question23;
                            continue;
                        case 23:
                            my_aql.Section3[i] = line.sec3.Question24;
                            my_aql.Section4[i] = line.sec4.Question24;
                            my_aql.Section6[i] = line.sec6.Question24;
                            my_aql.Section7[i] = line.sec7.Question24;
                            continue;
                        case 24:

                            my_aql.Section3[i] = line.sec3.Question25;
                            my_aql.Section4[i] = line.sec4.Question25;
                            my_aql.Section7[i] = line.sec7.Question25;

                            break;
                        default:
                            break;
                    }

                }
               
                my_aql.changeData();
                string my_data = my_aql.ToString();
                // Console.WriteLine(my_data);
                // Console.WriteLine($"chars = {my_data.Count()}");
                AQLs.Add(my_aql);
                count++;
                //  my_aql.Dump();
                //NBT = Convert.ToInt64(line.Person.Person.Person.Person.Person.Person.
                //Console.WriteLine($"{count} -> {line.Name}{line.Surname}{line.Question21}' '{line.Question22} {line.Question23} {line.Question24}  {line.Question25}{line.RefNo}");


            }

            var groupedData = AQLs
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
            
            var nbts = AQLs
                      .Select( n => n.RefNo.ToString()).ToList();

            string nbtFile = Path.Combine(Path.GetDirectoryName(filepath), "NBT_References.txt");
            File.WriteAllLines(nbtFile, nbts);

            ////Afrikaans test
            //var AQLA = AQLs
            //    .Where(t => t.Group.Contains("AKG"))
            //    .ToList();

            //string AQLA_test = AQLA.FirstOrDefault()?.Group;
            //string[] aqlaparts = AQLA_test.Trim().Split(' ');
            //string part1 = aqlaparts[1].Substring(0,2);
            //int testcode = Convert.ToInt32(aqlaparts[1].Substring(2));
            //string AQLAfileName = "AQLA" + testcode.ToString("D4") + part1 + testcode.ToString() + ".txt";

            ////English test
            //var AQLE = AQLs
            //.Where(t => t.Group.Contains("AQL"))
            //.ToList();

            //string AQLE_test = AQLE.FirstOrDefault()?.Group;
            //string[] aqleparts = AQLE_test.Trim().Split(' ');
            //string part10 = aqleparts[1].Substring(0, 2);
            //int Etestcode = Convert.ToInt32(aqlaparts[1].Substring(2));
            //string AQLEfileName = "AQLE" + Etestcode.ToString("D4") + part10 + Etestcode.ToString() + ".txt";

            //// Writing records to file
            //string directory = Path.GetDirectoryName(filepath);
            //string AQLEfile = Path.Combine(directory, AQLEfileName);
            //string AQLAfile = Path.Combine(directory, AQLAfileName);
            //using (StreamWriter writer = new StreamWriter(AQLEfile))
            //{
            //    foreach (var record in AQLE)
            //    {
            //        writer.WriteLine(record.ToString());
            //    }
            //}
            //using (StreamWriter writer = new StreamWriter(AQLAfile))
            //{
            //    foreach (var record in AQLA)
            //    {
            //        writer.WriteLine(record.ToString());
            //    }
            //}

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
                //string my_data = my_mat.ToString();
                //Console.WriteLine(my_data);

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
            //worksheet.Cell(currentRow, 4).Value = person.SURNAME;
            //worksheet.Cell(currentRow, 5).Value = person.FIRST_NAME;
            //worksheet.Cell(currentRow, 6).Value = person.INITIALS;
            //worksheet.Cell(currentRow, 1).Value = person.RefNo;
            //worksheet.Cell(currentRow, 3).Value = person.Barcode;
            //worksheet.Cell(currentRow, 4).Value = person.SURNAME;
            //worksheet.Cell(currentRow, 5).Value = person.FIRST_NAME;
            //worksheet.Cell(currentRow, 6).Value = person.INITIALS;
        }
        string directory5 = Path.GetDirectoryName(filepath);
        string excelfile = Path.Combine(directory5, "NBT_AnswershettBio.xlsx");
        workbook.SaveAs(excelfile);
    }

    Console.ReadLine();
}
