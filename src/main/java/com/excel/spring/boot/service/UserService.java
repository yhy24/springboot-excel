package com.excel.spring.boot.service;

import com.excel.spring.boot.util.User;

import java.util.List;

/**
 * @Author: yhy
 * @Date: 2018/7/31 10:39
 * @Version 1.0
 */
public interface UserService {

    public User selectUserById(Integer id);

    public List<User> selectUsers();

}
