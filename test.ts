import Automizer from 'pptx-automizer';
import PptxGenJS from 'pptxgenjs';
import { ISlide, modify, XmlElement } from 'pptx-automizer/dist';

( async () => {
    const automizer = new Automizer ( {
        templateDir: './input',
        outputDir  : './output',

        // Start clean so we can control ordering precisely
        removeExistingSlides: true,

        // Safer when mixing different templates
        autoImportSlideMasters: true,

        verbosity: 2
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

    const pptx = await pptxgenjs.write ( { outputType: 'arraybuffer' } ) as Buffer;

    const performancePptxLabel = 'performance';
    const rootPptxLabel = 'root';

    const ROOT_FILE = 'ac-terms-and-uf-terms.pptx';

    let pres = automizer
        .loadRoot ( ROOT_FILE )
        .load ( ROOT_FILE, rootPptxLabel )
        .load ( pptx, performancePptxLabel );

    // We don't need this prefix, using it bc the input I'm using has it as well, but we could simply have the "<DATE_PLACEHOLDER>"
    // directly in the sub header place which will ease this update for translations
    const subHeaderPrefix = '(the "Access Fund") |';
    const PLACEHOLDER = `${subHeaderPrefix} <DATE_PLACEHOLDER>`;

    const addRootReplacingByXml = ( n: number ) =>
        // pptx-automizer supports xmldom as callback
        pres.addSlide ( rootPptxLabel, n, async ( sl ) => {
            const ids = await sl.getAllTextElementIds ();
            for ( const id of ids ) {
                sl.modifyElement ( id, ( el: XmlElement ) => {
                    const texts = el.getElementsByTagName ( 'a:t' );
                    for ( let i = 0; i < texts.length; i++ ) {
                        const t = texts.item ( i );
                        if ( t?.textContent && t.textContent === PLACEHOLDER ) {
                            const currDate = new Date ();
                            const monthName = currDate.toLocaleString (
                                'en-US',
                                { month: 'long' }
                            );
                            const year = currDate.getFullYear ();

                            t.textContent = `${subHeaderPrefix} ${monthName} ${year}`;
                        }
                    }
                } );
            }
        } );

    // PS: The below could be done in a loop as well, doing it manually for clarity

    // Add first slide of the root presentation
    addRootReplacingByXml ( 1 );

    // Add the second slide of the root presentation
    addRootReplacingByXml ( 2 );

    // Add the single generated performance slide in the middle of the root presentation
    pres.addSlide ( performancePptxLabel, 1 );

    // Add the remaining slides of the root presentation
    addRootReplacingByXml ( 3 );
    addRootReplacingByXml ( 4 );
    addRootReplacingByXml ( 5 );

    await pres.write ( 'merged-final.pptx' );
} ) ();
