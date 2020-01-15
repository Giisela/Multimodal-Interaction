package scxmlgen.Modalities;

import scxmlgen.interfaces.IModality;

/**
 *
 * @author nunof
 */
public enum SecondMod implements IModality{

    /*RED("[color][RED]",1500),
    BLUE("[color][BLUE]",1500),
    YELLOW("[color][YELLOW]",1500);
    
    CHANGE_THEME("[6][theme]",1500),
    START_PRESENTATION("[presentation]", 1500),
    STOP_PRESENTATION("[presentation]", 1500),
    NEXT("[slide]",1500),
    NEXT_PRESENTATION("[slide]",1500),
    PREVIOUS("[slide]",1500),
    PREVIOUS_PRESENTATION("[slide]",1500),
    READ_TITLE("[read]",1500),
    READ_TEXT("[read]",1500),
    READ_NOTES("[read]",1500),
    OPEN_POWERPOINT("[openPowerPoint]",1500),
    CLOSE_POWERPOINT("[close]",1500),
    JUMPTO("[slide]",1500),
    JUMPTO_PRESENTATION("[slide]",1500),
    ZOOMOUT("[ZoomO][]",1500),
    ZOOMIN("[ZoomI][]",1500),
    CROPOUT("[CropO][]",1500),
    CROPIN("[CropI][]",1500),
    YES("[confirmation]",1500),
    NO("[confirmation]",1500);
    */
    
    
    CHANGE_THEME("[6][ThemaR]",1500),
    START_PRESENTATION("[4][Open]", 1500),
    STOP_PRESENTATION("[0][Close]", 1500),
    NEXT_GESTURES("[3][NextR]",1500),
    PREVIOUS_GESTURES("[5][PreviouL]",1500),
    NEXT("[NEXT]",1500),
    PREVIOUS("[PREVIOUS]",1500),
    
    OPEN_POWERPOINT("[OPEN_POWERPOINT]",1500),
    CLOSE_POWERPOINT("[CLOSE]",1500),
    
    ZOOMOUT("[8][ZoomO]",1500),
    ZOOMIN("[7][ZoomI]",1500),
    CROPOUT("[2][CropO]",1500),
    CROPIN("[1][CropI]",1500),

    
    YES("[YES]",1500),
    NO("[NO]",1500);
    ;
    
    private String event;
    private int timeout;


    SecondMod(String m, int time) {
        event=m;
        timeout=time;
    }

    @Override
    public int getTimeOut() {
        return timeout;
    }

    @Override
    public String getEventName() {
        //return getModalityName()+"."+event;
        return event;
    }

    @Override
    public String getEvName() {
        return getModalityName().toLowerCase()+event.toLowerCase();
    }
    
}
