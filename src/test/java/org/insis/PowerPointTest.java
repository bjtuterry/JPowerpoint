package org.insis;

import org.insis.openxml.powerpoint.PowerPoint;
import org.insis.openxml.powerpoint.PowerPointHelper;
import org.insis.openxml.powerpoint.Slide;
import org.insis.openxml.powerpoint.Text;
import org.insis.openxml.powerpoint.TextBox;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class PowerPointTest 
    extends TestCase
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public PowerPointTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( PowerPointTest.class );
    }

    /**
     * Rigourous Test :-)
     */
    public void testPowerPoint()
    {
        PowerPoint ppt = PowerPointHelper.create("D:\\test.pptx");
        Slide slide = ppt.addSlide();
        TextBox tb = slide.addTextBox(0, 0, 100, 100);
        Text text = tb.addText("test", false);
        text.setFontSize(24);
        ppt.save();
    }
}
