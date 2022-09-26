package org.example;

import java.util.UUID;

public class CMark {
    private UUID id;
    private double value;

    public double getValue() {
        return value;
    }

    public void setValue(double value) {
        this.value = value;
    }
    private CStudent student;

    public CStudent getStudent() {
        return student;
    }

    public void setStudent(CStudent student) {
        this.student = student;
    }

    public UUID getId() {
        return id;
    }

    public void setId(UUID id) {
        this.id = id;
    }

    @Override
    public String toString() {
        return "Студент: "+((student==null) ? "не указан":student.getName())+" Балл: "+value;
    }
}
