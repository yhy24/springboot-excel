package com.excel.spring.boot.dao;

import com.excel.spring.boot.util.User;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

/**
 * @Author: yhy
 * @Date: 2018/7/31 10:39
 * @Version 1.0
 */
@Mapper
public interface UserDao {
    public User selectUserId(Integer id);

    public void updateUserById(User user);

    public List<User> selectUsers();
}
