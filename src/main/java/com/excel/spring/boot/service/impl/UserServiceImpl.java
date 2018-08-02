package com.excel.spring.boot.service.impl;

import com.excel.spring.boot.dao.UserDao;
import com.excel.spring.boot.service.UserService;
import com.excel.spring.boot.util.User;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * @Author: yhy
 * @Date: 2018/7/31 10:40
 * @Version 1.0
 */
@Service
public class UserServiceImpl implements UserService{
    @Autowired
    UserDao userDao;
    @Override
    public User selectUserById(Integer id) {
        return userDao.selectUserId(id);
    }

    @Override
    public List<User> selectUsers() {
        return userDao.selectUsers();
    }
}
