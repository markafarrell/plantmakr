// (C) Copyright 2002-2005 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted, 
// provided that the above copyright notice appears in all copies and 
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting 
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC. 
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to 
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

//-----------------------------------------------------------------------------
//----- acrxEntryPoint.h
//-----------------------------------------------------------------------------
#include "StdAfx.h"
#include "resource.h"

//-----------------------------------------------------------------------------
#define szRDS _RXST("MAF")

//-----------------------------------------------------------------------------
//----- ObjectARX EntryPoint
class CPlantSchedulerApp : public AcRxArxApp {

public:
	CPlantSchedulerApp () : AcRxArxApp () {}

	virtual AcRx::AppRetCode On_kInitAppMsg (void *pkt) {
		// TODO: Load dependencies here

		// You *must* call On_kInitAppMsg here
		AcRx::AppRetCode retCode =AcRxArxApp::On_kInitAppMsg (pkt) ;
		
		// TODO: Add your initialization code here

		return (retCode) ;
	}

	virtual AcRx::AppRetCode On_kUnloadAppMsg (void *pkt) {
		// TODO: Add your code here

		// You *must* call On_kUnloadAppMsg here
		AcRx::AppRetCode retCode =AcRxArxApp::On_kUnloadAppMsg (pkt) ;

		// TODO: Unload dependencies here

		return (retCode) ;
	}

	virtual void RegisterServerComponents () {
	}

public:

	// - MAFPlantScheduler._PLANT command (do not rename)
	static void MAFPlantScheduler_PLANT(void)
	{
		ads_name sset; //the selection set
		long sslen = 0;
		wchar_t str[100];
		struct resbuf eb1; //vars for filtering the selection set
		wchar_t sbuf1[100];

		std::map<std::wstring, int> plants;

		std::map<std::wstring, plant> database;

		std::map<int, std::string> catagories;

		std::string filename = get_db_path();

		//load the database

		load_db(filename, &database, &catagories);

		// Create the filter
		
		eb1.restype = 0; //entity type
		wcscpy(sbuf1, _T("TEXT,MTEXT,MULTILEADER"));

		eb1.resval.rstring = sbuf1; //Select TEXT, MTEXT and MLEADER entities
		eb1.rbnext = NULL;

		// Create the selection set getting only MTEXT, TEXT and MLEADER
		if(acedSSGet(_T("P"), NULL, NULL, &eb1 , sset) == RTNORM)
		{
			acutPrintf(_T("Selection Set Created Successfully\n"));
		}
		
		if(acedSSLength(sset, &sslen) != RTNORM)
		{
			acedAlert(_T("Nothing preselected. Selecting everything\n"));
			acedSSGet(_T("A"), NULL, NULL, &eb1 , sset);
		}
		else
		{
			swprintf(str, _T("Items selected: %ld\n"), sslen);
			acedAlert(str);
		}

		extract_data(sset, &plants);

		out_to_excel(&plants, &database, &catagories);
		acedSSFree(sset);
		
	}

private: 

	static std::string get_db_path()
	{
		std::ifstream config;

		config.open("C:/Program Files/PlantMakr/plantmakr.cfg", std::ifstream::in);

		std::string line;

		getline(config, line);

		return line;
	}

	static std::string get_regex()
	{
		std::ifstream config;

		config.open("C:/Program Files/PlantMakr/plantmakr.cfg", std::ifstream::in);

		std::string line;

		getline(config, line);
		getline(config, line);

		return line;
	}

	static std::wstring StringToWString(const std::string& s)
	{
		std::wstring temp(s.length(),L' ');
		std::copy(s.begin(), s.end(), temp.begin());
		return temp;
	}

	static void load_db(std::string filename, std::map<std::wstring, plant> * database, std::map<int, std::string> * catagories)
	{
		std::ifstream plantdb;

		plantdb.open(filename.c_str(), std::ifstream::in);

		std::string line;
		std::string field;
		std::string temp_str;
		std::string plant_code_s;
		std::wstring plant_code_ws;
		wchar_t plant_code[100];
		wchar_t str[500];
		int num_cats;
		int i=0;

		if(plantdb.is_open())
		{
			getline(plantdb, line);
			
			sscanf(line.c_str(), "%d", &num_cats);
			
			for(i=0; i<num_cats; i++)
			{
				std::string cat;
				std::string cat_id_s;
				int cat_id;
				getline(plantdb, line);

				cat = get_next_word(&line);

				cat_id_s = get_next_word(&line);
				sscanf(cat_id_s.c_str(), "%d", &cat_id);
				
				(*catagories)[cat_id] = cat;
				
			}
			getline(plantdb, line); //get rid of column headings
			while(!plantdb.eof())
			{
				getline(plantdb, line);
				if(line.length() == 0)
				{
					break;
				}

				plant_code_s = get_next_word(&line);
				plant_code_ws = StringToWString(plant_code_s);

				//swprintf(str, _T("%s\n"), plant_code_ws.c_str());
				//acedAlert(str);

				(*database)[plant_code_ws].bio_name = get_next_word(&line);
			
				//swprintf(str, _T("%s\n"), StringToWString((*database)[plant_code_ws].bio_name).c_str());
				//acedAlert(str);

				(*database)[plant_code_ws].common_name = get_next_word(&line);

				//swprintf(str, _T("%s\n"), StringToWString((*database)[plant_code_ws].common_name).c_str());
				//acedAlert(str);

				sscanf(get_next_word(&line).c_str(), "%d", &((*database)[plant_code_ws].plant_type));

				//swprintf(str, _T("pot: %d\n"), (*database)[plant_code_ws].plant_type);
				//acedAlert(str);

				get_next_word(&line); //Mature Height
				get_next_word(&line); //Mature Width
				get_next_word(&line); //Australian Native
				get_next_word(&line); //Deciduous
				sscanf(get_next_word(&line).c_str(), "%d", &((*database)[plant_code_ws].pot_size));

				//swprintf(str, _T("%d\n"), (*database)[plant_code_ws].pot_size);
				//acedAlert(str);

				sscanf(get_next_word(&line).c_str(), "%d", &((*database)[plant_code_ws].install_spacing));

				//swprintf(str, _T("%d\n"), (*database)[plant_code_ws].install_spacing);
				//acedAlert(str);

				sscanf(get_next_word(&line).c_str(), "%d", &((*database)[plant_code_ws].install_height));

				//swprintf(str, _T("%d\n"), (*database)[plant_code_ws].install_height);
				//acedAlert(str);

//				swprintf(str, _T("%s\n%s\n%s\n%d\n%d\n%d\n%d\n"), plant_code_ws.c_str(), StringToWString((*database)[plant_code_ws].bio_name).c_str(), StringToWString((*database)[plant_code_ws].common_name).c_str(), 
//																(*database)[plant_code_ws].plant_type, (*database)[plant_code_ws].pot_size, 
//																(*database)[plant_code_ws].install_spacing, (*database)[plant_code_ws].install_height);
//				acedAlert(str);
			}
		}
		else
		{
			acedAlert(_T("Database not found"));
			exit(0);
		}
	}

	static std::string get_next_word(std::string * line)
	{
		int end_of_word=0;
		std::string next_word;
		end_of_word = line->find(",", 0);
		if(end_of_word == std::string::npos)
		{
			next_word = *line;
		}
		else
		{
			next_word = line->substr(0, end_of_word);
			*line = line->substr(end_of_word+1);
		}
		return next_word;
	}
	
	static void change_col_width(int col, float width, vole::object worksheet)
	{
		using vole::object;
		using vole::collection;

		object columns = worksheet.get_property<object>(L"Columns");
		object column = columns.get_property<object>(L"Item", col);

		column.put_property(L"ColumnWidth", width);
	}

	static void out_to_excel(std::map<std::wstring, int> * plants, std::map<std::wstring, plant> * database, std::map<int, std::string> * catagories)
	{
		using vole::object;
		using vole::collection;
		
		std::map<std::wstring, int>::iterator plant_it;
		std::map<int, std::string>::iterator catagory_it;

		int i=5;
		comstl::ole_init    coinit;
		wchar_t str[100];
		wchar_t range_str[10];
		//std::wstring str;
		std::wstring plant_code_ws;

		object excelApp = object::create("Excel.Application", CLSCTX_LOCAL_SERVER, vole::coercion_level::nonDestructiveCoercion);

		object workBooks = excelApp.get_property<object>(L"Workbooks");

		workBooks.invoke_method<void>(L"Add");

		object workBook = workBooks.get_property<object>(L"Item", 1);

		object windows = workBook.get_property<object>(L"Windows");

		object window = windows.get_property<object>(L"Item", 1);

		object worksheets = workBook.get_property<object>(L"Worksheets");
		object worksheet = worksheets.get_property<object>(L"Item", 1);
		
		window.put_property(L"DisplayGridlines", 0);
	 
		object titlerange = excelApp.get_property<vole::object>(L"Range", "B2");
		titlerange.put_property(L"Formula", "Plant Schedule");
		object tfont = titlerange.get_property<object>(L"Font");
		tfont.put_property(L"Size", 14);
		tfont.put_property(L"Bold", 1);
		
		change_col_width(2, 8.4, worksheet);
		change_col_width(3, 37.5, worksheet);
		change_col_width(4, 30, worksheet);
		change_col_width(5, 9.25, worksheet);
		change_col_width(6, 12, worksheet);
		change_col_width(7, 9.25, worksheet);
		change_col_width(8, 8.38, worksheet);

		vole::object range = excelApp.get_property<vole::object>(L"Range", "B4");
		range.put_property(L"Formula", "Code");
		
		vole::object range2 = excelApp.get_property<vole::object>(L"Range", "C4");
		range2.put_property(L"Formula", "Botanical Name");
		
		vole::object range3 = excelApp.get_property<vole::object>(L"Range", "D4");
		range3.put_property(L"Formula", "Common Name");
		
		vole::object range6 = excelApp.get_property<vole::object>(L"Range", "E4");
		range6.put_property(L"Formula", "Pot Size");
		
		vole::object range5 = excelApp.get_property<vole::object>(L"Range", "F4");
		range5.put_property(L"Formula", "No. per m2");

		vole::object range4 = excelApp.get_property<vole::object>(L"Range", "G4");
		range4.put_property(L"Formula", "Height");
		
		vole::object range1 = excelApp.get_property<vole::object>(L"Range", "H4");
		range1.put_property(L"Formula", "Quantity");
		
		object col_labels_range = excelApp.get_property<object>(L"Range", "B4:H4");
		object col_labels_font = col_labels_range.get_property<object>(L"Font");
		col_labels_font.put_property(L"Bold", 1);

		vole::object border = col_labels_range.get_property<vole::object>(L"Borders");
		border.put_property(L"LineStyle", 1);
	
		for(catagory_it = catagories->begin(); catagory_it !=catagories->end(); catagory_it++)
		{
			i++;
			swprintf(range_str, _T("B%d"), i);
			object range = excelApp.get_property<object>(L"Range", range_str);
			object cat_font = range.get_property<object>(L"Font");
			cat_font.put_property(L"Bold", 1);
			range.put_property(L"Formula", catagory_it->second.c_str());

			i++;
			
			for(plant_it = plants->begin(); plant_it != plants->end(); plant_it++)
			{
				if((*database).find(plant_it->first) != (*database).end())
				{
					if((*database)[plant_it->first].plant_type == catagory_it->first)
					{
						plant_code_ws = plant_it->first;

						//swprintf(str, _T("%s\n\n%s\n%d"), StringToWString(catagory_it->second).c_str(), plant_code_ws.c_str(), plant_it->second);
					
						//acedAlert(str);	

						//CODE
						swprintf(range_str, _T("B%d"), i);
						object range1 = excelApp.get_property<object>(L"Range", range_str);
						range1.put_property(L"Formula", plant_it->first.c_str());					
						//Botanical Name
						swprintf(range_str, _T("C%d"), i);
						object range3 = excelApp.get_property<object>(L"Range", range_str);
						range3.put_property(L"Formula", (*database)[plant_it->first].bio_name.c_str());
						//Common Name
						swprintf(range_str, _T("D%d"), i);
						object range4 = excelApp.get_property<object>(L"Range", range_str);
						range4.put_property(L"Formula", (*database)[plant_it->first].common_name.c_str());
						//Height
						swprintf(range_str, _T("G%d"), i);
						object range5 = excelApp.get_property<object>(L"Range", range_str);
						range5.put_property(L"Formula", (*database)[plant_it->first].install_height);
						//Spacing
						swprintf(range_str, _T("F%d"), i);
						object range6 = excelApp.get_property<object>(L"Range", range_str);
						range6.put_property(L"Formula", (*database)[plant_it->first].install_spacing);
						//Pot Size
						swprintf(range_str, _T("E%d"), i);
						object range7 = excelApp.get_property<object>(L"Range", range_str);
						range7.put_property(L"Formula", (*database)[plant_it->first].pot_size);
						//Quantity
						swprintf(range_str, _T("H%d"), i);
						object range2 = excelApp.get_property<object>(L"Range", range_str);
						range2.put_property(L"Formula", plant_it->second);

						swprintf(range_str, _T("B%d:H%d"), i, i);
						object plant_range = excelApp.get_property<object>(L"Range", range_str);
						object plant_border = plant_range.get_property<object>(L"Borders");

						plant_border.put_property(L"LineStyle", 1);

						i++;
					}
				}
				
			}
		}

		//output plants not in database

		i++;

		swprintf(range_str, _T("B%d"), i);
		object unknown_range = excelApp.get_property<object>(L"Range", range_str);
		object cat_font = unknown_range.get_property<object>(L"Font");
		cat_font.put_property(L"Bold", 1);
		unknown_range.put_property(L"Formula", "Unknown");
		i++;

		for(plant_it = plants->begin(); plant_it != plants->end(); plant_it++)
		{
			if((*database).find(plant_it->first) == (*database).end()) //plant is not in database
			{
				plant_code_ws = plant_it->first;
				//swprintf(str, _T("%s\n\n%s\n%d"), StringToWString(catagory_it->second).c_str(), plant_code_ws.c_str(), plant_it->second);
				
				//acedAlert(str);	

				//CODE
				swprintf(range_str, _T("B%d"), i);
				object range1 = excelApp.get_property<object>(L"Range", range_str);
				range1.put_property(L"Formula", plant_it->first.c_str());					
				//Quantity
				swprintf(range_str, _T("H%d"), i);
				object range2 = excelApp.get_property<object>(L"Range", range_str);
				range2.put_property(L"Formula", plant_it->second);

				swprintf(range_str, _T("B%d:H%d"), i, i);
				object plant_range = excelApp.get_property<object>(L"Range", range_str);
				object plant_border = plant_range.get_property<object>(L"Borders");

				plant_border.put_property(L"LineStyle", 1);

				i++;
			}
		}

		swprintf(range_str, _T("B4:H%d"), i);

		object whole_range = excelApp.get_property<object>(L"Range", range_str);
		object whole_border = whole_range.get_property<object>(L"Borders");
		
		object bottom_border = whole_border.get_property<object>(L"Item", 9); //bottom edge
		object top_border = whole_border.get_property<object>(L"Item", 8); //top edge
		object left_border = whole_border.get_property<object>(L"Item", 7); //left edge
		object right_border = whole_border.get_property<object>(L"Item", 10); //right edge

		object whole_font = whole_range.get_property<object>(L"Font");
	
		bottom_border.put_property(L"LineStyle", 1); //continuous
		bottom_border.put_property(L"Weight", -4138); //medium

		top_border.put_property(L"LineStyle", 1); //continuous
		top_border.put_property(L"Weight", -4138); //medium
		
		left_border.put_property(L"LineStyle", 1); //continuous
		left_border.put_property(L"Weight", -4138); //medium
		
		right_border.put_property(L"LineStyle", 1); //continuous
		right_border.put_property(L"Weight", -4138); //medium

		whole_font.put_property(L"Name", "HelveticaNeue LT 47 LightCn");

		excelApp.put_property(L"Visible", true);
	}

	static void extract_data(ads_name sset, std::map<std::wstring, int> * plants)
	{
		long sslen;
		int i;
		ads_name entname;
		AcDbEntity * pEnt;
		AcDbMLeader * ml_ent;
		AcDbMText * mt_ent;
		AcDbText * t_ent;

		boost::wregex expression(StringToWString(get_regex()));
		boost::match_results<const wchar_t *> what;

		const ACHAR * ent_type;
		ACHAR * text;

		int quantity=0;

		char str[100];
		wchar_t wstr[100];

		acedSSLength(sset, &sslen);

		for(i=0; i<sslen; i++)
		{	
			if(acedSSName(sset, i, entname) != RTNORM)
			{
				continue;
			}			
			pEnt = openEntity(entname, AcDb::kForRead);

			ent_type = pEnt->isA()->name();

			if(wcscmp(ent_type, _T("AcDbMLeader"))==0)
			{
				ml_ent = (AcDbMLeader *)pEnt;
				mt_ent = ml_ent->mtext();
				text = mt_ent->text();
			}	
			else if(wcscmp(ent_type, _T("AcDbMText")) == 0)
			{				
				mt_ent = (AcDbMText *)pEnt;
				text = mt_ent->text();
			}
			else if(wcscmp(ent_type, _T("AcDbText"))==0)
			{
				t_ent = (AcDbText *)pEnt;
				text = t_ent->textString();
			}
			else
			{
				pEnt->close();
				continue;
			}
			pEnt->close();

			text = (wchar_t *)text;

			if(boost::regex_search((const wchar_t *)text, what, expression))
			{
				quantity = 0;
				quantity = wcstol(what.str(1).c_str(), (wchar_t **)(what.str(1).c_str() + sizeof(wchar_t)*wcslen(what.str(1).c_str())), 10);
				
				quantity = quantity + (*plants)[what.str(2).c_str()];

				(*plants)[what.str(2).c_str()] = quantity;

				//wsprintf(wstr, _T("%s: %d\n"), what.str(2).c_str(), quantity);

				//acedAlert(wstr);
			}
		}
		return;
	}
	
	static AcDbEntity*
	openEntity(ads_name entity, AcDb::OpenMode openMode)
	{
		AcDbObjectId eId;
		// Exchange the ads_name for an object ID.
		//
		acdbGetObjectId(eId, entity);

		AcDbEntity * pEnt;
		acdbOpenObject(pEnt, eId, openMode);

		return pEnt;
	}
} ;

//-----------------------------------------------------------------------------
IMPLEMENT_ARX_ENTRYPOINT(CPlantSchedulerApp)

ACED_ARXCOMMAND_ENTRY_AUTO(CPlantSchedulerApp, MAFPlantScheduler, _PLANT, PLANT, ACRX_CMD_TRANSPARENT, NULL)