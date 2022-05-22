/**
 * Create a sleep prediction spreadsheet
 *
 * @copyright Copyright 2020-2021 Sleep Diary Authors <sleepdiary@pileofstuff.org>
 *
 * @license
 *
 * Permission is hereby granted, free of charge, to any person
 * obtaining a copy of this software and associated documentation
 * files (the "Software"), to deal in the Software without
 * restriction, including without limitation the rights to use, copy,
 * modify, merge, publish, distribute, sublicense, and/or sell copies
 * of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
 * BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
 * ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
 * CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

add_export("prediction_spreadsheet",function( workbook, statistics ) {

    const day_length = (statistics.summary_days||{}).average || 24*ONE_HOUR,
          time = new Date().getTime() - 86400000,
          asleep_at = (
              time
              - ( time % 86400000 ) // reset to midnight GMT
              + statistics.schedule.sleep.average // average day length
              + statistics.schedule.sleep.durations.length/2 * ( day_length - 24*ONE_HOUR ) // skew
          ),
          awake_at  = (
              time
              - ( time % 86400000 ) // reset to midnight GMT
              + statistics.schedule.wake .average + // average day length
              + statistics.schedule.wake.durations.length/2 * ( day_length - 24*ONE_HOUR ) // skew
              + ( // make sure sleep comes before wake
                  statistics.schedule.wake.average < statistics.schedule.sleep.average
              ) * 86400000
          ),
          estimate_interval = ONE_HOUR/2,
          one_day           = 24*ONE_HOUR,
          worksheet         = workbook.addWorksheet("Estimates"),
          settings = [
              [ "Average over this many days", statistics.schedule.sleep.timestamps.length-1 ],
              [ "Base uncertainty", estimate_interval / one_day ],
              [ "Daily uncertainty multiplier", 1.1 ],
              [ "Last recorded sleep", { formula: "=MAX(A:A)" } ],
              [ "Start of sleep-averaging period", { formula: "=VLOOKUP(J6-J3,A:A,1)" } ],
              [ "Average time between sleeps", { formula: "=(J6-J7)/J3" } ],
              [ "Last recorded wake", { formula: "=MAX(B:B)" } ],
              [ "Start of wake-averaging period", { formula: "=VLOOKUP(J9-J3,B:B,1)" } ],
              [ "Average time between wakes", { formula: "=(J9-J10)/J3" } ],
              [ "Average day length", { formula: "=(J8+J11)/2" } ],
          ],
          day_format = "ddd\\ MMM\\ D,\\ HH:MM",
          sleep_column = {
              width: 18,
              style: {
                  numFmt: day_format,
                  font: {
                      name: 'Calibri',
                      color: { argb: "FFFFFFFF" }
                  },
                  fill: {
                      type: "pattern",
                      pattern: "solid",
                      fgColor: {argb:"FF000040"},
                  },
              }
          },
          wake_column = {
              width: 18,
              style: {
                  numFmt: day_format,
                  font: {
                      name: 'Calibri',
                      color: { argb: "FF000000" }
                  },
                  fill: {
                      type: "pattern",
                      pattern: "solid",
                      fgColor: {argb:"FFFFFFA0"},
                  },
              }
          },

          sleep_prediction_column = {
              width: 18,
              style: {
                  numFmt: day_format,
                  font: {
                      name: 'Calibri',
                      color: { argb: "FFFFFFFF" }
                  },
                  fill: {
                      type: "pattern",
                      pattern: "solid",
                      fgColor: {argb:"FF000080"},
                  },
              }
          },
          wake_prediction_column = {
              width: 18,
              style: {
                  numFmt: day_format,
                  font: {
                      name: 'Calibri',
                      color: { argb: "FF000000" }
                  },
                  fill: {
                      type: "pattern",
                      pattern: "solid",
                      fgColor: {argb:"FFFFFFC0"},
                  },
              }
          },

          heading_style = {
              font: {
                  name: 'Calibri',
                  bold: true,
              },
              fill: {
                  type: "pattern",
                  pattern: "solid",
                  fgColor: {argb:"FFEEEEEE"},
              },
              alignment: {
                  vertical: "middle",
                  horizontal: "center"
              },
          },
          rows = [
              ["Observed times",undefined,undefined,"Estimated sleep time",undefined,"Estimated wake time",undefined,undefined,"Algorithm"],
              ["Asleep at","Awake at",undefined,"Earliest","Latest","Earliest","Latest",undefined,"Setting","Value"]
          ].concat(
              statistics.schedule.sleep.timestamps.map( (v,n) => [
                  new Date(v),
                  new Date(statistics.schedule.wake.timestamps[n]),
                  undefined
              ])
          ),
          max_rows = rows.length+Math.min( 14, statistics.schedule.sleep.timestamps.length-1 );
    ;

    for ( let n=2; n<max_rows; ++n ) {
        const row = rows[n] = rows[n] || [undefined,undefined,undefined];
        row[3] = { formula: "=IF(A"+(n+1)+",A"+(n+1)+",D"+n+"+$J$12-$J$4*$J$5)" };
        row[4] = { formula: "=IF(A"+(n+1)+",A"+(n+1)+",E"+n+"+$J$12+$J$4*$J$5)" };
        row[5] = { formula: "=IF(B"+(n+1)+",B"+(n+1)+",F"+n+"+$J$12-$J$4*$J$5)" };
        row[6] = { formula: "=IF(B"+(n+1)+",B"+(n+1)+",G"+n+"+$J$12+$J$4*$J$5)" };
    }

    // add settings:
    settings.forEach( (v,n) => {
        rows[n+2][8] = v[0];
        rows[n+2][9] = v[1];
    });

    // set column styles:
    worksheet.columns = [
        sleep_column,
        wake_column,
        {},
        sleep_prediction_column, sleep_prediction_column,
        wake_prediction_column,  wake_prediction_column,
        {},
        { width: 18, style: { numFmt: day_format } },
        { width: 18, style: { numFmt: day_format } },
    ];

    worksheet.addRows(rows);

    // set heading styles:
    worksheet.mergeCells("A1:B1");
    worksheet.mergeCells("D1:E1");
    worksheet.mergeCells("F1:G1");
    worksheet.mergeCells("I1:J1");
    [ 'A', 'B', 'D', 'E', 'F', 'G', 'I', 'J' ].forEach(
        column => [ 1, 2 ].forEach(
            row => worksheet.getCell( column + row ).style = heading_style
        )
    );

    // set setting formats:
    worksheet.getCell( 'J3' ).style = { numFmt: "0" };
    worksheet.getCell( 'J5' ).style = { numFmt: "0.00" };
    [6,7].forEach( cell => worksheet.getCell( 'J'+cell ).style = sleep_column.style );
    [9,10].forEach( cell => worksheet.getCell( 'J'+cell ).style =  wake_column.style );
    [4,8,11,12].forEach( cell => worksheet.getCell( 'J'+cell ).style = { numFmt: "[HH]:MM" } );

    return workbook;

});
