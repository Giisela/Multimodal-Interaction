/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package genfusionscxml;

import java.io.IOException;
import scxmlgen.Fusion.FusionGenerator;
import scxmlgen.Modalities.Output;
import scxmlgen.Modalities.Speech;
import scxmlgen.Modalities.SecondMod;

/**
 *
 * @author nunof
 */
public class GenFusionSCXML {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {

    FusionGenerator fg = new FusionGenerator();
  
    /**
    fg.Sequence(Speech.SQUARE, SecondMod.RED, Output.SQUARE_RED);
    fg.Sequence(Speech.SQUARE, SecondMod.BLUE, Output.SQUARE_BLUE);
    fg.Sequence(Speech.SQUARE, SecondMod.YELLOW, Output.SQUARE_YELLOW);
    fg.Complementary(Speech.TRIANGLE, SecondMod.RED, Output.TRIANGLE_RED);
    fg.Complementary(Speech.TRIANGLE, SecondMod.BLUE, Output.TRIANGLE_BLUE);
    fg.Complementary(Speech.TRIANGLE, SecondMod.YELLOW, Output.TRIANGLE_YELLOW);
    fg.Complementary(Speech.CIRCLE, SecondMod.RED, Output.CIRCLE_RED);
    fg.Complementary(Speech.CIRCLE, SecondMod.BLUE, Output.CIRCLE_BLUE);
    fg.Complementary(Speech.CIRCLE, SecondMod.YELLOW, Output.CIRCLE_YELLOW);
    
    //fg.Single(Speech.CIRCLE, Output.CIRCLE);  //EXAMPLE
    //fg.Redundancy(Speech.CIRCLE, SecondMod.CIRCLE, Output.CIRCLE);  //EXAMPLE

    // Redundancy
    fg.Redundancy(Speech.NEXT, SecondMod.NEXT, Output.NEXT);
    fg.Redundancy(Speech.NEXT_PRESENTATION, SecondMod.NEXT_PRESENTATION, Output.NEXT_PRESENTATION);
    fg.Redundancy(Speech.PREVIOUS, SecondMod.PREVIOUS, Output.PREVIOUS);
    fg.Redundancy(Speech.OPEN_POWERPOINT, SecondMod.OPEN_POWERPOINT, Output.OPEN_POWERPOINT);
    fg.Redundancy(Speech.CLOSE_POWERPOINT, SecondMod.CLOSE_POWERPOINT, Output.CLOSE_POWERPOINT);
    fg.Redundancy(Speech.JUMPTO, SecondMod.JUMPTO, Output.JUMPTO);
    fg.Redundancy(Speech.JUMPTO_PRESENTATION, SecondMod.JUMPTO_PRESENTATION, Output.JUMPTO_PRESENTATION);
    fg.Redundancy(Speech.READ_TITLE, SecondMod.READ_TITLE, Output.READ_TITLE);
    fg.Redundancy(Speech.READ_TEXT, SecondMod.READ_TEXT, Output.READ_TEXT);
    fg.Redundancy(Speech.READ_NOTES,SecondMod.READ_NOTES,Output.READ_NOTES);
    fg.Redundancy(Speech.ZOOMOUT, SecondMod.ZOOMOUT, Output.ZOOMOUT);
    fg.Redundancy(Speech.ZOOMIN, SecondMod.ZOOMIN, Output.ZOOMIN);
    fg.Redundancy(Speech.CROPOUT,SecondMod.CROPOUT, Output.CROPOUT);
    fg.Redundancy(Speech.CROPIN, SecondMod.CROPIN, Output.CROPIN);
    fg.Redundancy(Speech.YES, SecondMod.YES, Output.YES);
    fg.Redundancy(Speech.NO, SecondMod.NO, Output.NO);


    // Complementarity
    fg.Complementary(SecondMod.CHANGE_THEME, Speech.THEME_ONE, Output.CHANGE_THEME1);
    fg.Complementary(SecondMod.CHANGE_THEME, Speech.THEME_TWO, Output.CHANGE_THEME2);
    fg.Complementary(SecondMod.CHANGE_THEME, Speech.THEME_THREE, Output.CHANGE_THEME3);
    fg.Complementary(SecondMod.START_PRESENTATION, Speech.YES, Output.START_PRESENTATION_YES);
    fg.Complementary(SecondMod.START_PRESENTATION, Speech.NO, Output.START_PRESENTATION_NO);
    fg.Complementary(SecondMod.STOP_PRESENTATION, Speech.YES, Output.STOP_PRESENTATION_YES);
    fg.Complementary(SecondMod.STOP_PRESENTATION, Speech.NO, Output.STOP_PRESENTATION_NO);

    */
    //Single
    fg.Single(Speech.JUMPTO, Output.JUMPTO);
    fg.Single(Speech.JUMPTO_PRESENTATION, Output.JUMPTO_PRESENTATION);
    
    fg.Single(Speech.READ_TITLE, Output.READ_TITLE);
    fg.Single(Speech.READ_TEXT,  Output.READ_TEXT);
    fg.Single(Speech.READ_NOTES, Output.READ_NOTES);
    
    fg.Single(Speech.READ_TITLE_PRESENTATION, Output.READ_TITLE_PRESENTATION);
    fg.Single(Speech.READ_TEXT_PRESENTATION,  Output.READ_TEXT_PRESENTATION);
    fg.Single(Speech.READ_NOTES_PRESENTATION, Output.READ_NOTES_PRESENTATION);
    
    fg.Single(SecondMod.ZOOMOUT, Output.ZOOMOUT);
    fg.Single(SecondMod.ZOOMIN, Output.ZOOMIN);
    
    fg.Single(SecondMod.CROPOUT, Output.CROPOUT);
    fg.Single(SecondMod.CROPIN, Output.CROPIN);
    
    fg.Single(Speech.OPEN_POWERPOINT, Output.OPEN_POWERPOINT);
    fg.Single(Speech.CLOSE_POWERPOINT, Output.CLOSE_POWERPOINT);
    
    
    
    // Redundancy
    fg.Redundancy(Speech.NEXT, SecondMod.NEXT_GESTURES, Output.NEXT);
    fg.Redundancy(Speech.NEXT_PRESENTATION, SecondMod.NEXT_GESTURES, Output.NEXT_PRESENTATION);
    fg.Redundancy(Speech.NEXT_GESTURES, SecondMod.NEXT_GESTURES, Output.NEXT);
    fg.Redundancy(Speech.PREVIOUS, SecondMod.PREVIOUS_GESTURES, Output.PREVIOUS);
    fg.Redundancy(Speech.PREVIOUS_PRESENTATION, SecondMod.PREVIOUS_GESTURES, Output.PREVIOUS_PRESENTATION);
    fg.Redundancy(Speech.PREVIOUS_GESTURES, SecondMod.PREVIOUS_GESTURES, Output.PREVIOUS);
    fg.Redundancy(Speech.START_PRESENTATION, SecondMod.START_PRESENTATION,  Output.START_PRESENTATION);
    fg.Redundancy(Speech.STOP_PRESENTATION, SecondMod.STOP_PRESENTATION, Output.STOP_PRESENTATION);
    
   
    // Complementarity
    fg.Complementary(Speech.CHANGE_THEME_ONE, SecondMod.CHANGE_THEME, Output.CHANGE_THEME_ONE);
    fg.Complementary(Speech.CHANGE_THEME_TWO, SecondMod.CHANGE_THEME, Output.CHANGE_THEME_TWO);
    fg.Complementary(Speech.CHANGE_THEME_THREE, SecondMod.CHANGE_THEME, Output.CHANGE_THEME_THREE);
    

    
    fg.Build("fusion.scxml");
        
        
    }
    
}
