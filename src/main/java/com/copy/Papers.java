package com.copy;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Papers
{
    private String name = "";

    private List<String> key_value = new ArrayList<String>();

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<String> getKey_value() {
        return key_value;
    }

    public void setKey_value(List<String> key_value) {
        this.key_value = key_value;
    }

    @Override
    public String toString() {
        return "Papers [name=" + name + ", key_value=" + Arrays.toString(key_value.toArray()) + "]";
    }


}
