/*
 * Copyright (c) 2022
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package nedis;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;

public class Launcher {
    
    static {
        ZipSecureFile.setMinInflateRatio(0.0);
    }

    public static void main(String[] args) throws Exception {
        boolean trim = true;
        String file1 = "/home/nedis/Downloads/AR_267_03102022143425.xlsx";
        String file2 = "/home/nedis/Downloads/AR_266_03102022175129.xlsx";

        try (Workbook workbook1 = WorkbookFactory.create(new File(file1));
             Workbook workbook2 = WorkbookFactory.create(new File(file2))) {
            new WorkbookComparator(trim, workbook1, workbook2).compare();
        }
    }
}
