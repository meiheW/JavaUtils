package myutils;

import javafx.scene.input.DataFormat;

import java.sql.Timestamp;
import java.time.LocalDateTime;
import java.util.Date;

public class DateUtil {


    public static Timestamp Now(){
        return Timestamp.valueOf(LocalDateTime.now());
    }

    public static Date now(){
        return new Date();
    }



}
