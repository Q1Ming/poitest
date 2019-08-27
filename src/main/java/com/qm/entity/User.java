package com.qm.entity;

import lombok.Data;
import lombok.experimental.Accessors;

import java.util.Date;

/**
 * @author Qm
 * 测试导出数据用的实体类，没有任何意义
 **/
@Data
@Accessors(chain = true)
public class User {
    private String id;
    private String name;
    private Date bir;
}
