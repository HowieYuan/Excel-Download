package com.howie.excelandfile;

import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA
 *
 * @Author yuanhaoyue swithaoy@gmail.com
 * @Description 生成人员数据
 * @Date 2018-03-22
 * @Time 21:32
 */
@Service
public class PersonFactroy {
    public List<Person> getPersonList() {
        List<Person> list = new ArrayList<>();
        list.add(new Person(1, "Howie", 20, "female"));
        list.add(new Person(2, "Wade", 25, "male"));
        list.add(new Person(3, "Duncan", 30, "male"));
        list.add(new Person(4, "Kobe", 35, "male"));
        list.add(new Person(5, "James", 40, "male"));
        return list;
    }
}
