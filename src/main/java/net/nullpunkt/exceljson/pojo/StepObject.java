package net.nullpunkt.exceljson.pojo;

import java.util.ArrayList;
import java.util.Collection;

public class StepObject{

    private ArrayList<HeaderObject> steps = new ArrayList<HeaderObject>();

    public void addStepRow(HeaderObject row) {
        steps.add(row);
    }

    public void fillColumns() {
        for(HeaderObject tmp: steps) {
        }
    }

    // GET/SET

    public ArrayList<HeaderObject> getSteps() {
        return steps;
    }

    public void setSteps(ArrayList<HeaderObject> steps) {
        this.steps = steps;
    }
}

