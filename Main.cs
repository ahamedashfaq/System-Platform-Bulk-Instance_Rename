//------------------------------------------------------------------------------
//AUTHORS: MACRO-INTEGRATION PTE LTD

using System;
using System.IO;
using System.Text;
using ArchestrA.GRAccess;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Collections;

class tagRename
{
	[STAThread]
	static void Main()

	{
        GRAccessApp grAccess = new GRAccessAppClass();
        Console.WriteLine("GR Access Initiated... ");

        Console.WriteLine("Enter NODE name TO CONTINUE:");
        string nodeName = Console.ReadLine(); //"MAC-02
        
        //string nodeName = Environment.MachineName;
        // Type your grname and press enter
        
        Console.WriteLine("Enter GR name TO CONTINUE:");
        string galaxyName = Console.ReadLine(); //"F10A2GR01MIGRATE

        Console.WriteLine();
        Console.WriteLine("GR Name is: " + galaxyName);
        Console.WriteLine();
        Console.WriteLine("Processing.... Please wait......");
       
        IGalaxies gals = grAccess.QueryGalaxies(nodeName);
        ICommandResult cmd;
        IGalaxy galaxy = gals[galaxyName];
        galaxy = grAccess.QueryGalaxies(nodeName)[galaxyName];
		// log in
		galaxy.Login( "ADMINISTRATOR", "" );
        cmd = galaxy.CommandResult;
        if (!cmd.Successful)
        {
            Console.WriteLine("Login to galaxy Failed :" +
                            cmd.Text + " : " +
                            cmd.CustomMessage);
            return;
        }
        else
        {
            Console.WriteLine(galaxyName + " login successfull");
            Console.WriteLine();
        }

		// get the taglist template
        string oldtaglist = @"c:\temp\oldtaglist.txt";
        string newtaglist = @"c:\temp\newtaglist.txt";

        string [] oldInst = File.ReadLines(oldtaglist).ToArray();
        string [] newInst = File.ReadLines(newtaglist).ToArray();

        var oldTlineCount = File.ReadAllLines(oldtaglist).Length;
        var newTlineCount = File.ReadAllLines(newtaglist).Length;

        if (oldTlineCount==newTlineCount)
        {
            Console.WriteLine("Both file tag count matched successfully");
            Console.WriteLine();
        }
        else
        {
            Console.WriteLine(" Match unsuccessfull");
            Console.WriteLine();
        }





        Console.WriteLine("No.of tags in new Tag File: " + newTlineCount);
        Console.WriteLine();

        Console.WriteLine("Renaming Started");
        Console.WriteLine();

        for (int i = 0; i < oldTlineCount; i++)
        {
            
            string oldInstanceArray = oldInst[i];
            Console.Write("Processing:" + " " + i + " " + oldInst[i]);
            Console.WriteLine();
            
            string[] refobj = new string[] { oldInstanceArray };

  
            IgObjects queryResult = galaxy.QueryObjectsByName(
                EgObjectIsTemplateOrInstance.gObjectIsInstance,
                ref refobj);

            IInstance sampleinst = (IInstance)queryResult[1];

            sampleinst.Tagname = newInst[i];
            Console.Write(oldInst[i] + " changed to "+ newInst[i]);
            Console.WriteLine();
        }



        galaxy.Logout();

        Console.WriteLine();
        Console.Write("Press ENTER to quit: ");
        string dummy;
        dummy = Console.ReadLine();
		
	}
}
