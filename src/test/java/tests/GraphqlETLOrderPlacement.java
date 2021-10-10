package tests;

import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;
import utils.DBUtil;
import utils.DataProviderUtil;
import utils.ExcelUtil;

import java.sql.SQLException;

public class GraphqlETLOrderPlacement {

    int count = 1;
    boolean isFirstRow = true;

        @BeforeClass
    public void beforeStepsForStock() throws SQLException {
        DBUtil.beforeStart(DataProviderUtil.getSKUList("AE"));
    }


    @Test(dataProvider = "getDataAE", dataProviderClass = DataProviderUtil.class)
    public void placeOrders(String[] sku, String[] qty, String streetAddress, String city, String phone, String discount, String fname, String lname, String countryCode, String cat) throws Exception {

            ExcelUtil.updategsheet(count,"orderID");
    }

}
