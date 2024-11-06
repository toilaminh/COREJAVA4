package org.example;

public class Emp {
    private static int index = 0;
    private int EMP_INDEX;
    private String ID;
    private String NAME;
    public Emp(String id, String name){
        EMP_INDEX = index;
        ID = id;
        NAME = name;
        index += 1;
    }

}
