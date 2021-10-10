package utils;

import java.sql.*;
import java.util.ArrayList;
import java.util.Objects;

public class DBUtil {

    public static void beforeStart(ArrayList<ArrayList<String[]>> arrayList) throws SQLException {

        ArrayList<String> selectQueries = new ArrayList<>();
        ArrayList<String> updateQueries = new ArrayList<>();
        ArrayList<String> stockTableUpdateQueries = new ArrayList<>();

        for (ArrayList<String[]> a : arrayList){
            for (String[] f : a){
                for (int i =0; i < f.length; i++){
                    selectQueries.add("SELECT ");
                    updateQueries.add("UPDATE");
                }
            }
        }


        Connection con = null;
        boolean isUpdateRequired = false;
        boolean isStockTableUpdateRequired = false;

        try {
            con = DriverManager.getConnection("jdbc:mysql://domain:port/db","username","pw");

            for(int i=0;i<selectQueries.size();i++){

                Statement stmt1 = con.createStatement();
                System.out.println(selectQueries.get(i));
                ResultSet rs= stmt1.executeQuery(selectQueries.get(i));
                while (rs.next()){
                    System.out.println("QTY available : "+rs.getInt(1));
                    if (rs.getInt(1) < 50){
                        System.out.println("Before update qty : "+rs.getInt(1));
                        isUpdateRequired = true;
                        if (rs.getInt(1) <= 0){
                            isStockTableUpdateRequired = true;
                        }
                    }
                }
                rs.close();
                if (isUpdateRequired){
                    Statement stmt2 = con.createStatement();
                    System.out.println(updateQueries.get(i));
                    stmt2.execute(updateQueries.get(i));
                    System.out.println("Updated Qty");
                    if (isStockTableUpdateRequired){
                        Statement stmt3 = con.createStatement();
                        System.out.println(stockTableUpdateQueries.get(i));
                        stmt3.execute(stockTableUpdateQueries.get(i));
                        System.out.println("Updated  stock table");
                    }
                }
            }
        }
        catch (SQLException e) {
            e.printStackTrace();
        }
        Objects.requireNonNull(con).close();
    }
}
