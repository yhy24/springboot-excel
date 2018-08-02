package com.excel.spring.boot.util;

import java.io.Serializable;

public class User implements Serializable{
    private static final long serialVersionUID = -5046972860468900599L;

    private Integer id;
    private String username;
    private String password;
    private String intest;
    private String phone;
    private String email;
    private Integer age;
    private String mark;
    private String sexly;
    private String code;
    private String gender;
    private String indicSeq;
    private String create_date;
    private String modify_date;


    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }
    public String getCreate_date() {
        return create_date;
    }

    public String getModify_date() {
        return modify_date;
    }

    public void setCreate_date(String create_date) {

        this.create_date = create_date;
    }

    public void setModify_date(String modify_date) {
        this.modify_date = modify_date;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public User() {
    }
    public User(Integer id, String username, String password, String intest, String phone, String email, Integer age, String mark, String sexly) {
        this.id = id;
        this.username = username;
        this.password = password;
        this.intest = intest;
        this.phone = phone;
        this.email = email;
        this.age = age;
        this.mark = mark;
        this.sexly = sexly;
    }

    public Integer getId() {
        return id;
    }

    public String getUsername() {
        return username;
    }

    public String getPassword() {
        return password;
    }


    public String getPhone() {
        return phone;
    }

    public String getEmail() {
        return email;
    }

    public Integer getAge() {
        return age;
    }

    public String getMark() {
        return mark;
    }

    public String getSexly() {
        return sexly;
    }

    public void setId(Integer id) {

        this.id = id;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public void setMark(String mark) {
        this.mark = mark;
    }

    public void setSexly(String sexly) {
        this.sexly = sexly;
    }

    public String getIntest() {
        return intest;
    }

    public void setIntest(String intest) {
        this.intest = intest;
    }

    public String getIndicSeq() {
        return indicSeq;
    }

    public void setIndicSeq(String indicSeq) {
        this.indicSeq = indicSeq;
    }

    @Override
    public String toString() {
        return "User{" +
                "id=" + id +
                ", username='" + username + '\'' +
                ", password='" + password + '\'' +
                ", intest='" + intest + '\'' +
                ", phone='" + phone + '\'' +
                ", email='" + email + '\'' +
                ", age=" + age +
                ", mark='" + mark + '\'' +
                ", sexly='" + sexly + '\'' +
                ", code='" + code + '\'' +
                ", gender='" + gender + '\'' +
                ", indicSeq='" + indicSeq + '\'' +
                ", create_date='" + create_date + '\'' +
                ", modify_date='" + modify_date + '\'' +
                '}';
    }
}
