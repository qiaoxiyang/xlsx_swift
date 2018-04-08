//
//  ViewController.swift
//  xlsx_swift
//
//  Created by xiyang on 2018/3/30.
//  Copyright © 2018年 xiyang. All rights reserved.
//

import Cocoa

class ViewController: NSViewController {

    override func viewDidLoad() {
        super.viewDidLoad()

        // Do any additional setup after loading the view.
        
        let filePath = Bundle.main.path(forResource: "test", ofType: "xlsx")
        
        let spreadSheet = BRAOfficeDocumentPackage.open(filePath)
        let workSheet = spreadSheet?.workbook.worksheets.first as! BRAWorksheet
        let formula = workSheet.cell(forCellReference: "A1").stringValue()
        
        print(formula)
        
        
    }

    override var representedObject: Any? {
        didSet {
        // Update the view, if already loaded.
        }
    }


}

