	using System;
	using System.Windows.Forms;
	using System.Xml;
	using System.IO;
	using System.Diagnostics;
	using Eplan.EplApi.Scripting;
	using Eplan.EplApi.ApplicationFramework;
	using Eplan.EplApi.Base;
	using Eplan.EplApi.Gui;
public class OpenFileClass
{	
	[DeclareMenu]
	public void CreateMenuFunction()
	{
		Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
		uint menuid ;
		menuid = oMenu.AddMainMenu("Project And Macro", Eplan.EplApi.Gui.Menu.MainMenuName.eMainMenuProject ,"Open Project...","ElexcoOpen","Open menu with an opendialog",1);
		oMenu.AddMenuItem("Open Macro Project...", "ElexcoOpenMacro", "",menuid,1,false,false);
		oMenu.AddMenuItem("Open Current Project Folder...", "ElexcoOpenProjectFolder", "",menuid,1,false,false);
		oMenu.AddMenuItem("Setting...", "ElexcoSetFolder", "",menuid,0,true,true);
		return;
	}

    [DeclareAction("ElexcoOpen")]
    public void Function()
    {
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();


        OpenFileDialog openFileDialog1 = new OpenFileDialog();
		if (oSettings.ExistSetting("USER.SCRIPTS.OpenProject.PathFolderProject")){
			openFileDialog1.InitialDirectory = oSettings.GetStringSetting("USER.SCRIPTS.OpenProject.PathFolderProject", 0);
		}
        openFileDialog1.Filter = "Eplan Project (*.elk)|*.elk|All files (*.*)|*.*" ;
        openFileDialog1.RestoreDirectory = true ;

        if(openFileDialog1.ShowDialog() == DialogResult.OK)
	    {
            OpenProject(openFileDialog1.FileName.ToString()); // Kommentar
	    }      
        return;
    }

    [DeclareAction("ElexcoOpenMacro")]
    public void MacroFunction()
    {
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

        OpenFileDialog openFileDialog1 = new OpenFileDialog();
        if (oSettings.ExistSetting("USER.SCRIPTS.OpenProject.PathFolderMacro")){
			openFileDialog1.InitialDirectory = oSettings.GetStringSetting("USER.SCRIPTS.OpenProject.PathFolderMacro", 0);
		}
        openFileDialog1.Filter = "Eplan Macro Project (*.elk)|*.elk|All files (*.*)|*.*" ;
        openFileDialog1.RestoreDirectory = true ;

        if(openFileDialog1.ShowDialog() == DialogResult.OK)
	    {
            OpenProject(openFileDialog1.FileName.ToString()); // Kommentar
	    }      
        return;
    }

	[DeclareAction("ElexcoOpenProjectFolder")]
    public void OpenProjectFolderFunction()
    {
		Process.Start(PathMap.SubstitutePath("$(PROJECTPATH)"));
        return;
    }
	
	[DeclareAction("ElexcoSetFolder")]
    public void SetDefaultFolder()
    {
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

		if (!oSettings.ExistSetting("USER.SCRIPTS.OpenProject.PathFolderProject"))
		{
			oSettings.AddStringSetting("USER.SCRIPTS.OpenProject.PathFolderProject",
			new string[] { },
			new string[] { }, "test setting",
			new string[] { "Default value of test setting" },
			ISettings.CreationFlag.Insert);	
		}
		if (!oSettings.ExistSetting("USER.SCRIPTS.OpenProject.PathFolderMacro"))
		{
			oSettings.AddStringSetting("USER.SCRIPTS.OpenProject.PathFolderMacro",
			new string[] { },
			new string[] { }, "test setting",
			new string[] { "Default value of test setting" },
			ISettings.CreationFlag.Insert);	
		}
		MessageBox.Show("You need to select your project base folder.");

		FolderBrowserDialog  folderBrowserDialog1 = new FolderBrowserDialog ();
		if (folderBrowserDialog1.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog1.SelectedPath))
		{
			oSettings.SetStringSetting("USER.SCRIPTS.OpenProject.PathFolderProject", folderBrowserDialog1.SelectedPath, 0);					
		}
		//FolderBrowserDialog  folderBrowserDialog1 = new FolderBrowserDialog ();
		if (folderBrowserDialog1.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowserDialog1.SelectedPath))
		{
			oSettings.SetStringSetting("USER.SCRIPTS.OpenProject.PathFolderMacro", folderBrowserDialog1.SelectedPath, 0);					
		}

        return;
    }

    [DeclareMenu]
    public void MenuFunction()
    {
        Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
        uint menuid ;
        menuid = oMenu.AddMainMenu("Elexco", Eplan.EplApi.Gui.Menu.MainMenuName.eMainMenuHelp,"Open Project...","ElexcoOpen","Open menu with an opendialog",1);
        return;
    }
    
    public static void OpenProject(string projectname)
    {
        CommandLineInterpreter oCLI = new CommandLineInterpreter();
        ActionCallingContext acc = new ActionCallingContext();
        acc.AddParameter("PROJECTNAME", projectname);
        oCLI.Execute("EDIT", acc);
    }
    
    
}
