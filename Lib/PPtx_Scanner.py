import collections
import collections.abc
from pptx import *

def PPtx_Scanner(Path = "Res/TEST.pptx"):
    FileName = Path
    prs = Presentation(FileName)

    SlidesNo = len(prs.slides) #45
    slide_With_Chart = []
    print("Presentation have "+str(SlidesNo)+" slides.\n")

    for slideNo in range(0,SlidesNo):
        #print("\n\nSlide : "+str(slide)+" #########\n\n")
        slide = prs.slides[slideNo]
        for shape in range(0,len(slide.shapes)):
            if slide.shapes[shape].has_chart:
                slide_With_Chart.append(str(slideNo)+":"+str(shape))
                break

    print(slide_With_Chart)
    print("\n\n\n")

    for No in slide_With_Chart:
        print(No)
        slide = prs.slides[int(No.split(":")[0])]
        chart = slide.shapes[int(No.split(":")[1])].chart
        for plot in range(0,len(chart.plots)):
            print("Plot : " + str(plot))
            for ser in range(0,len(chart.plots[plot].series)):
                print("Serie : " + str(ser))
                print(chart.plots[plot].series[ser].values)
        print("\n\n\n")

def PPtx_Page(Path = "Res/TEST.pptx"):
    FileName = Path
    prs = Presentation(FileName)

    SlidesNo = len(prs.slides) #45
    return SlidesNo

def PPtx_GetText(Path,slide,shape,para,run):
    FileName = Path
    prs = Presentation(FileName)
    slide = prs.slides[slide]
    shape = slide.shapes[shape]
    para = shape.text_frame.paragraphs[para]
    run = para.runs[run]
    text = run.text
    return text

def PPtx_TextFrame(Path = "Res/TEST.pptx",slide=0):
    FileName = Path
    prs = Presentation(FileName)
    slideNo = slide
    slide = prs.slides[slide] #45
    TextFrames = []
    for shapeNo in range(0,len(slide.shapes)):
        if slide.shapes[shapeNo].has_text_frame:
            #print("Shape : "+str(shapeNo))
            shape = slide.shapes[shapeNo]
            for paraNo in range(0,len(shape.text_frame.paragraphs)):
                #print("Para : "+str(paraNo))
                para = shape.text_frame.paragraphs[paraNo]
                for runNo in range(0,len(para.runs)):
                    #print("Run : "+str(runNo))
                    #print(para.runs[runNo].text)
                    TextFrames.append(str(slideNo)+":"+str(shapeNo)+":"+str(paraNo)+":"+str(runNo))
        #print("\n")
    return (TextFrames)

if __name__ == "__main__":
    print(PPtx_GetText("Res/TEST.pptx",1,1,0,0))