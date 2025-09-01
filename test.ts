import Automizer from 'pptx-automizer';
import PptxGenJS from 'pptxgenjs';
import { modify } from 'pptx-automizer/dist';

( async () => {
    const automizer = new Automizer ( {
        templateDir: './input',
        outputDir  : './output',

        // Start clean so we can control ordering precisely
        removeExistingSlides: true,

        // Safer when mixing different templates
        autoImportSlideMasters: true
    } );

    const pptxgenjs = new PptxGenJS ();
    const slide = pptxgenjs.addSlide ();

    slide.addText (
        [
            {
                text   : 'Slide created from PptxGenJS, to merge with pptx-automizer slides',
                options: { breakLine: true, paraSpaceAfter: 10 }
            },
            {
                text   : 'This would be the performance slides',
                options: { breakLine: true }
            }
        ],
        {
            x       : 1,
            y       : 5,
            fontSize: 15
        }
    );

    const pptx = await pptxgenjs.write ( { outputType: 'nodebuffer' } ) as Buffer;

    const performancePptxLabel = 'performance';
    const rootPptxLabel = 'root';

    const ROOT_FILE = 'ac-terms-and-uf-terms.pptx';

    let pres = automizer
        .loadRoot ( ROOT_FILE )
        .load ( ROOT_FILE, rootPptxLabel )
        .load ( pptx, performancePptxLabel );

    // @TODO: modify the <DATE_PLACEHOLDER> in the root presentation slid

    // PS: The below could be done in a loop as well, doing it manually for clarity

    // Add first slide of the root presentation
    pres.addSlide ( rootPptxLabel, 1 );

    // Add the second slide of the root presentation
    pres.addSlide ( rootPptxLabel, 2 );

    // Add the single generated performance slide in the middle of the root presentation
    pres.addSlide ( performancePptxLabel, 1 );

    // Add the remaining slides of the root presentation
    pres.addSlide ( rootPptxLabel, 3 );
    pres.addSlide ( rootPptxLabel, 4 );
    pres.addSlide ( rootPptxLabel, 5 );

    await pres.write ( 'merged-final.pptx' );
} ) ();
