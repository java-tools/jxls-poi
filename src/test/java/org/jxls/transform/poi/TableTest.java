package org.jxls.transform.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

public class TableTest {

    @Test
    public void testTable() throws IOException {
        List<TableTestObject> list = new ArrayList<TableTestObject>();
        for (int i = 1; i <= 100; i++) {
            list.add(new TableTestObject("name-" + i, "address-" + i));
        }
        Context ctx = new Context();
        ctx.putVar("list", list);
        
        InputStream in = TableTest.class.getResourceAsStream("table.xlsx"); // XLSX file with a table, jx:area and jx:each
        try {
            File dir = new File("test-output");
            dir.mkdirs();
            FileOutputStream out = new FileOutputStream(new File(dir, "table.xlsx"));
            try {
                Transformer transformer = JxlsHelper.getInstance().createTransformer(in, out);
                JxlsHelper.getInstance().processTemplate(ctx, transformer);
            } finally {
                out.close();
            }
        } finally {
            in.close();
        }
    }
    
    public static class TableTestObject {
        private final String name;
        private final String address;

        public TableTestObject(String name, String address) {
            this.name = name;
            this.address = address;
        }

        public String getName() {
            return name;
        }

        public String getAddress() {
            return address;
        }
    }
}
