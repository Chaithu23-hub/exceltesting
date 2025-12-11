package com.springexceltest.demo.entity;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;


@Entity
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Student {
    @Id
    private Long Id;
    @Column(nullable = false)
    private String Name;
    private String Gender;
    private Integer Age;
    private String Section;
    private Integer Science;
    private Integer English;
    private Integer History;
    private Integer Maths;

}
