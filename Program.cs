// See https://aka.ms/new-console-template for more information
using Cetap_Classes;
using Dumpify;
using System.Collections.ObjectModel;


Console.WriteLine("Territorium File reader");
string filepath = "D:/test/CEA_AQL_May_11.csv";
Collection<Read_online_AQL> my_recs;
List<onlineAQL> AQLs = new List<onlineAQL>();
try
{
    Console.WriteLine(filepath);
    Loadrecs reader = new Loadrecs(filepath,"AQL");
    my_recs = new Collection<Read_online_AQL>();
    my_recs = reader.AQLrecs;
    var section1 = my_recs  
   .Where(t => t.Test.Contains("Section 1")||t.Test.Contains("Afdeling 1"))
    .Select(t => new
    {
        t.RefNo, t.Test,t.Surname,t.Name, t.Group, t.Period, t.DOT, t.Email, t.StartTest,
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
        t.Question15,t.Question16,
        t.Question17
    })
   .ToList();

    var section2 = my_recs  
   .Where(t => t.Test.Contains("Section 2")||t.Test.Contains("Afdeling 2"))
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
        t.Question15,t.Question16,
        t.Question17
    })
   .ToList();

    var section3 = my_recs  
    .Where(t => t.Test.Contains("Section 3")||t.Test.Contains("Afdeling 3"))
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
         t.Question15,t.Question16,
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
    .Where(t => t.Test.Contains("Section 4")||t.Test.Contains("Afdeling 4"))
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
         t.Question15,t.Question16,
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
    .Where(t => t.Test.Contains("Section 5")||t.Test.Contains("Afdeling 5"))
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
         t.Question15,t.Question16,
         t.Question17
     })
    .ToList();

    var section6 = my_recs
    .Where(t => t.Test.Contains("Section 6")||t.Test.Contains("Afdeling 6"))
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
         t.Question15,t.Question16,
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
    .Where(t => t.Test.Contains("Section 7")||t.Test.Contains("Afdeling 7"))
    .Select(t => new { t.RefNo, t.Question1, t.Question2, t.Question3, t.Question4, t.Question5, t.Question6, t.Question7, t.Question8, t.Question9,
                       t.Question10, t.Question11, t.Question12, t.Question13, t.Question14, t.Question15,t.Question16, t.Question17, t.Question18, t.Question19, t.Question20})
    .ToList();

   // join sections into one record using unique refNo
    var AQL = section1
        .Join(section2, s1 => s1.RefNo, s2 => s2.RefNo, (s1, s2) => new { Person = s1, sec2 = s2})
        .Join(section3, s1 => s1.Person.RefNo, s3 => s3.RefNo, (s1, s3) => new { Person = s1, sec3 = s3})

        .Join(section4, s1 => s1.Person.Person.RefNo, s4 => s4.RefNo, (s1, s4) => new { Person = s1, sec4 = s4})

        .Join(section5, s1 => s1.Person.Person.Person.RefNo, s5 => s5.RefNo, (s1, s5) => new { Person = s1, sec5 = s5})

        .Join(section6, s1 => s1.Person.Person.Person.Person.RefNo, s6 => s6.RefNo, (s1, s6) => new { Person = s1, sec6 = s6})

        .Join(section7, s1 => s1.Person.Person.Person.Person.Person.RefNo, s7 => s7.RefNo, (s1, s7) => new { Person = s1, sec7 = s7})

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
                     continue;
                case 21:
                    my_aql.Section3[i] = line.sec3.Question22;
                    my_aql.Section4[i] = line.sec4.Question22;
                    my_aql.Section6[i] = line.sec6.Question22;
                    continue;
                case 22:
                    my_aql.Section3[i] = line.sec3.Question23;
                    my_aql.Section4[i] = line.sec4.Question23;
                    my_aql.Section6[i] = line.sec6.Question23;
                    continue;
                case 23:
                    my_aql.Section3[i] = line.sec3.Question24;
                    my_aql.Section4[i] = line.sec4.Question24;
                    my_aql.Section6[i] = line.sec6.Question24;
                    continue;
                case 24:
                   
                    my_aql.Section3[i] = line.sec3.Question25;
                    my_aql.Section4[i] = line.sec4.Question25;
                  
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


    var AQLA = AQLs
        .Where(t => t.Group.Contains("AKG"))
        .ToList();

    var AQLE = AQLs
    .Where(t => t.Group.Contains("AQL"))
    .ToList();


    // Writing records to file
    string directory = Path.GetDirectoryName(filepath);
    string AQLEfile = Path.Combine(directory, "AQLE.txt");
    string AQLAfile = Path.Combine(directory, "AQLA.txt");
    using (StreamWriter writer = new StreamWriter(AQLEfile))
    {
        foreach( var record in AQLE) 
        {
            writer.WriteLine(record.ToString());
        }
    }
    using (StreamWriter writer = new StreamWriter(AQLAfile))
    {
        foreach (var record in AQLA)
        {
            writer.WriteLine(record.ToString());
        }
    }

}
catch(FileNotFoundException)
{
    Console.WriteLine($"File '{filepath}' not found.");
}
Console.ReadLine();
