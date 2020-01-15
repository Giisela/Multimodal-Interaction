package scxmlgen.Modalities;

import scxmlgen.interfaces.IOutput;



public enum Output implements IOutput{

    /*SQUARE_RED("[shape][SQUARE][color][RED]"),
    SQUARE_BLUE("[shape][SQUARE][color][BLUE]"),
    SQUARE_YELLOW("[shape][SQUARE][color][YELLOW]"),
    TRIANGLE_RED("[shape][TRIANGLE][color][RED]"),
    TRIANGLE_BLUE("[shape][TRIANGLE][color][BLUE]"),
    TRIANGLE_YELLOW("[shape][TRIANGLE][color][YELLOW]"),
    CIRCLE_RED("[shape][CIRCLE][color][RED]"),
    CIRCLE_BLUE("[shape][CIRCLE][color][BLUE]"),
    CIRCLE_YELLOW("[shape][CIRCLE][color][YELLOW]"),
    CIRCLE("[shape][CIRCLE]");
    
    //SPEECH
    NEXT("[slide][NEXT]"),
    NEXT_PRESENTATION("[slide][NEXT_PRESENTATION]"),
    PREVIOUS("[slide][PREVIOUS]"),
    PREVIOUS_PRESENTATION("[slide][PREVIOUS_PRESENTATION]"),
    OPEN_POWERPOINT("[openPowerPoint][OPEN_POWERPOINT]"),
    CLOSE_POWERPOINT("[close][CLOSE]"),
    JUMPTO("[slide][JUMP_TO]"),
    JUMPTO_PRESENTATION("[slide][JUMP_TO_SLIDE_PRESENTATION]"),
    READ_TITLE("[read][TITLE_PRESENTATION]"),
    READ_TEXT("[read][TEXT_PRESENTATION]"),
    READ_NOTES("[read][NOTE_PRESENTATION]"),
    THEME_ONE("[theme][1]"),
    THEME_TWO("[theme][2]"),
    THEME_THREE("[theme][3]"),
    YES("[confirmation][YES]"),
    NO("[confirmation][NO]"),

    //GESTURES
    ZOOMOUT("[ZoomO][]"),
    ZOOMIN("[ZoomI][]"),
    CROPOUT("[CropO][]"),
    CROPIN("[CropI][]"),
    CHANGE_THEME1("[6][ThemeR][theme][1]"),
    CHANGE_THEME2("[6][themeR][theme][2]"),
    CHANGE_THEME3("[6][themeR][theme][3]"),
    START_PRESENTATION_YES("[presentation][START][confirmation][YES]"),
    START_PRESENTATION_NO("[presentation][START][confirmation][NO]"),
    STOP_PRESENTATION_YES("[presentation][STOP_PRESENTATION][confirmation][YES]"),
    STOP_PRESENTATION_NO("[presentation][STOP_PRESENTATION][confirmation][NO]");
    */
    
    
    //Speech
    
    OPEN_POWERPOINT("[openPowerPoint][OPEN_POWERPOINT]"),
    CLOSE_POWERPOINT("[close][CLOSE]"),
    JUMPTO("[slide][JUMP_TO]"),

    NEXT("[3][NextR][slide][NEXT]"),
    PREVIOUS("[5][PreviouL][slide][PREVIOUS]"),
    NEXT_PRESENTATION("[3][NextR][slide][NEXT_PRESENTATION]"),
    PREVIOUS_PRESENTATION("[5][PreviouL][slide][PREVIOUS_PRESENTATION]"),
    JUMPTO_PRESENTATION("[slide][JUMP_TO_SLIDE_PRESENTATION]"),
    
    READ_TITLE("[read][TITLE]"),
    READ_TEXT("[read][TEXT]"),
    READ_NOTES("[read][NOTE]"),
    
    READ_TITLE_PRESENTATION("[read][TITLE_PRESENTATION]"),
    READ_TEXT_PRESENTATION("[read][TEXT_PRESENTATION]"),
    READ_NOTES_PRESENTATION("[read][NOTE_PRESENTATION]"),
    
    //Gestures
    CHANGE_THEME_ONE("[6][ThemaR][theme][1]"),
    CHANGE_THEME_TWO("[6][ThemaR][theme][2]"),
    CHANGE_THEME_THREE("[6][ThemaR][theme][3]"),

    START_PRESENTATION("[4][Open][presentation][START]"),
    STOP_PRESENTATION("[0][Close][presentation][STOP_PRESENTATION]"),

    
    
    ZOOMOUT("[8][ZoomO]"),
    ZOOMIN("[7][ZoomI]"),
    CROPOUT("[2][CropO]"),
    CROPIN("[1][CropI]"),
    
    ;
    
    
    
    private String event;

    Output(String m) {
        event=m;
    }
    
    public String getEvent(){
        return this.toString();
    }

    public String getEventName(){
        return event;
    }
}
