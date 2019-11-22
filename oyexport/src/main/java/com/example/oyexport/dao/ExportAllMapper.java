package com.example.oyexport.dao;


import java.util.List;
import java.util.Map;

public interface ExportAllMapper {
    //跟进率
    List<Map<String, Object>> selectFollowRate(Map<String, Object> params);

    List<Map<String, Object>> selectFollowRateExt(Map<String, Object> params);

    //...其他
    List<Map<String, Object>> selectFollowRate2(Map<String, Object> params);

    List<Map<String, Object>> selectFollowRateExt2(Map<String, Object> params);


}