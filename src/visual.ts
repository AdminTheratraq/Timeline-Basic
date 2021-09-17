/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/
"use strict";
import "regenerator-runtime/runtime";
import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import DataView = powerbi.DataView;
import IVisualEventService = powerbi.extensibility.IVisualEventService;
import VisualObjectInstanceEnumerationObject = powerbi.VisualObjectInstanceEnumerationObject;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import DataViewObjects = powerbi.DataViewObjects;
import * as sanitizeHtml from "sanitize-html";
import * as d3 from 'd3';
import { VisualSettings } from "./settings";
import * as validDataUrl from "valid-data-url";

export interface TimelineData {
    Company: String;
    Type: string;
    Description: string;
    CompanyLink: string;
    Date: Date;
    HeaderImage: string;
    FooterImage: string;
    selectionId: powerbi.visuals.ISelectionId;
}

export class Visual implements IVisual {
    private target: d3.Selection<HTMLElement, any, any, any>;
    private header: d3.Selection<HTMLElement, any, any, any>;
    private footer: d3.Selection<HTMLElement, any, any, any>;
    private svg: d3.Selection<SVGElement, any, any, any>;
    private tooltip: d3.Selection<HTMLElement, any, any, any>;
    private margin = { top: 50, right: 40, bottom: 50, left: 40 };
    private settings: VisualSettings;
    private host: IVisualHost;
    private xScale: d3.ScaleTime<number, number>;
    private yScale: d3.ScaleLinear<number, number>;
    private colorDataByYear: any[];
    private initLoad = false;
    private events: IVisualEventService;
    private selectionManager: ISelectionManager;

    constructor(options: VisualConstructorOptions) {
        this.target = d3.select(options.element);
        this.header = d3.select(options.element).append("div");
        this.footer = d3.select(options.element).append("div");
        this.svg = d3.select(options.element).append('svg');
        this.tooltip = d3.select(options.element).append("div");
        this.events = options.host.eventService;
        this.host = options.host;
        this.selectionManager = options.host.createSelectionManager();
    }

    public update(options: VisualUpdateOptions) {
        this.events.renderingStarted(options);
        this.settings = Visual.parseSettings(options && options.dataViews && options.dataViews[0]);
        this.svg.selectAll('*').remove();
        this.header.selectAll("img").remove();
        this.header.classed("header", false);
        this.footer.selectAll("img").remove();
        this.footer.classed("footer", false);

        let vpWidth = options.viewport.width;
        let vpHeight = options.viewport.height;

        if (this.settings.dataPoint.layout.toLowerCase() === "header" || this.settings.dataPoint.layout.toLowerCase() === "footer") {
            vpHeight = options.viewport.height - 105;
        }

        var _this = this;
        this.svg.attr('height', vpHeight);
        this.svg.attr('width', vpWidth);

        let gHeight = vpHeight - this.margin.top - this.margin.bottom;
        let gWidth = vpWidth - this.margin.left - this.margin.right;

        this.target.on("contextmenu", () => {
            const mouseEvent: MouseEvent = <MouseEvent>d3.event;
            const eventTarget: any = mouseEvent.target;
            let dataPoint: any = d3.select(eventTarget).datum();
            this.selectionManager.showContextMenu(
                dataPoint ? dataPoint.selectionId : {},
                {
                    x: mouseEvent.clientX,
                    y: mouseEvent.clientY,
                }
            );
            mouseEvent.preventDefault();
        });

        var timelineData = Visual.CONVERTER(options.dataViews[0], this.host);
        timelineData = timelineData.slice(0, 100);
        let minDate, maxDate, currentDate;
        let timelineLocalData: TimelineData[] = [];
        currentDate = new Date();

        if (timelineData.length > 0) {
            minDate = new Date(currentDate.getFullYear() - this.settings.displayYears.PreviousYear, 0, 1);
            timelineLocalData = timelineData.map<TimelineData>((d) => { if (d.Date.getFullYear() >= minDate.getFullYear()) { return d; } }).filter(e => e);
            maxDate = new Date(currentDate.getFullYear() + this.settings.displayYears.FutureYear, 0, 1);
            timelineLocalData = timelineLocalData.map<TimelineData>((d) => { if (d.Date.getFullYear() <= maxDate.getFullYear()) { return d; } }).filter(e => e);
        }

        if (timelineLocalData.length > 0) {
            timelineData = timelineLocalData;
        } else if (timelineLocalData.length == 0) {
            minDate = new Date(Math.min.apply(null, timelineData.map(d => d.Date)));
            maxDate = new Date(Math.max.apply(null, timelineData.map(d => d.Date)));
            minDate = new Date(minDate.getFullYear(), 0, 1);
            maxDate = new Date(maxDate.getFullYear() + 1, 0, 1);
        }
        this.getColorDataByYear(minDate, maxDate);
        this.renderHeaderAndFooter(timelineData);
        this.renderXandYAxis(minDate, maxDate, gWidth, gHeight);
        this.renderTitle(vpWidth, gWidth);
        this.renderXAxisCirclesAndQuarters();
        this.renderTimeRangeLines(gHeight, timelineData);
        this.renderBox(timelineData, gWidth, gHeight);

        this.handleHyperLinkClick();

        this.svg.append('rect')
            .attr('class', 'border-rect')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', vpWidth)
            .attr('height', vpHeight + 10)
            .attr('stroke-width', '2px')
            .attr('stroke', '#333')
            .attr('fill', 'transparent');
        this.events.renderingFinished(options);
    }

    private handleHyperLinkClick() {
        let _this = this;
        let baseurl = "https://strategicanalysisinc.sharepoint.com";
        this.svg.selectAll("foreignObject a").on("click", function (e: Event) {
            e = e || window.event;
            let target: any = e.target || e.srcElement;
            let link = d3.select(this).attr("href");
            if (link.indexOf("http") === -1 || link.indexOf("http") > 0) {
                link = baseurl + link;
            }
            _this.host.launchUrl(link);
            d3.event.preventDefault();
            return false;
        });
    }

    private renderBox(timelineData: TimelineData[], gWidth, gHeight) {
        let _self = this;
        var gbox = this.svg.selectAll(".box").data(timelineData).enter().append("g")
            .attr('class', (d, i) => {
                if (d.Type === 'Regulatory') { return 'rect regulatory' }
                else if (d.Type === 'Commercial') { return 'rect commercial' }
                else if (d.Type === 'Clinical Trails') { return 'rect clinical-trails' }
            })
            .attr("title", (d) => { return sanitizeHtml(d.Description) + '(' + d.Date + ')'; })
            .attr("width", () => { return 100; })
            .attr("height", () => {
                return 70;
            })
            .attr('fill', '#ffffff')
            .attr('transform', (d, i) => {
                let y;
                if ((i % 2) === 0) {
                    let count = i / 2;
                    if (count % 2 === 0) { y = _self.yScale(-100); } 
                    else { y = _self.yScale(-60); }
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) { y = _self.yScale(90); } 
                    else { y = _self.yScale(40); }
                }
                return 'translate(' + (_self.xScale(d.Date) - 25) + ' ' + y + ')';
            });
        gbox.append("circle").attr('cx', 45).attr('cy', 0).attr('r', 7)
            .attr('fill', (d, i) => {
                var colorData = _self.colorDataByYear.find(c => c.year === d.Date.getFullYear());
                return colorData.color;
            })
        gbox.append("foreignObject")
            .html((d) => {
                let company = '<div><a href="' + sanitizeHtml(d.CompanyLink) + 
                '" style="color:' + "#000000" + ';">' + (d.Company ? sanitizeHtml(d.Company) : "") + "</a></div>";
                return ("<div>" + company + sanitizeHtml(d.Description) + "</div>");
            })
            .attr('x', '20')
            .attr('y', (d, i) => { if (i % 2 === 0) { return -5; } else { return -70; }
            })
            .attr("width", 100).attr("height", 60)
            .attr('fill', '#000000').attr('transform', 'translate(0,20)')
            .attr("font-size", 9);

        gbox.append("svg:image")
            .attr('xlink:href', (d) => {
                var icon = '';
                if (d.Type === 'Clinical Trials') {
                    icon = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABoAAAAiCAMAAAEBugmqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACoUExURQAAAFlZWVlZWVhYWFlZWVhYWFlZWVhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFhYWFhYWFpaWlVVVVlZWVlZWVlZWVhYWFlZWVdXV1hYWFhYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVhYWFlZWVpaWldXV1hYWFlZWVhYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVZWVrcVYpkAAAA4dFJOUwCLeJPn74A0JPejBDz/q7O7VEwYy2TT22wgh+N8CI/rKBCXn/MwQKf7rzhItwxQv8fPYNcUcN8sDjo1cQAAAAlwSFlzAAAXEQAAFxEByibzPwAAAU1JREFUKFOdku1WwjAMhoMwZEJEvhxzMqabQ5EhLTDv/85M0jBg7ofH55yleZe3bU5bEBCgSwEpIIW2tT5A5GpwpI8M1lpY5ZRbj+XQTWHOI0CRUjiSHrH1GuzryMHQTqKYmBan5aEIH0oDSen3QrH1yIRozls4VI0QFie1xHF/s6ak0anUlZBGB5dAGPrlAEw5DQuALB9RQwTm8xV7d9JLwn3pTKFRbUWcFN5cqO02e/o6KToCHx87lROe+bAr5VAnM22sVdSVL8cqtF71bzFL070Xx50gTWcv+vNw/87Lj1USn7LfcAKZTDTRURey8yDge4wzfi7CIVT0ILvykhr5a6kf4JumxFXJmI8NJMuBU1XpO893Bm5xHZU4o1u6KFF/Exo29JAWrsd6G8meDUK9dEFV4rdbQXfyj1krLXl3nPzm1YOi1VHzFct28QOudyFsm5wOugAAAABJRU5ErkJggg==";
                }
                else if (d.Type === 'Commercial') {
                    icon = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACIAAAAYCAMAAAHff++tAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADYUExURQAAAFlZWVlZWVhYWFtbW1lZWVhYWFlZWVhYWFhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFhYWFhYWFpaWlhYWFhYWFpaWllZWVhYWFhYWFhYWFVVVVlZWVlZWVlZWVdXV1hYWFlZWVdXV1hYWFhYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVlZWVhYWFlZWVpaWldXV1hYWFlZWT8/P1hYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFhYWFlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVlZWVZWVmNebiYAAABIdFJOUwCLeJMc5++AmzQk96MEPP+rs/ThRLtUTOnD8VwYy2TTdNtsIIfjfAiP6ygQl8Wf8zBAp/sIrzhItwxQv9rHWM9oYNcUcIPfLHjl7rgAAAAJcEhZcwAAFxEAABcRAcom8z8AAAGySURBVChTbZGLepNAEIUnVZKadjlqDbGa5loaSBUSVkoo5tYS8v5v5MyySj71/4CdOXNghl26JaIZAXLRdZpzKgAn0O4tNEfI+DYO5jIVIyU3hIpXZdwMl2Kpc1BcFSNrT7FiN0rd9eWtMeDPRYATFdVm4fnJkuWR/bargS3fQq38eCLO9/dGoeDkuG1gjWgF/3U951f3dwhB8zbbqSdzGK73n/pBq469evlDytZeMqoSSv2S86ffvWLXLAajXHk87tTkVmHHYebPjFAr6Tt1AWjVKAbOc54UdHLuh7xjWmMe4lQ8ywS4cahE/oLlGgcRqGJrTJvsiI4nMzHtL9DhI6IlXmtBcGinfBvzbKXswP9R4/hsHEZ7i8KGzvug8+aDCc8sseuaA/ybxhL0zdlc5urRdYxiaSzbB6BaUsb/Hmd53cNw1qh3WA3NKNKupQ72wEEbfsQPz51mDD2lKZfvBosJZ2F9on2OiPJbViZKOm0G3RfpZ6qGEPgJfIycMmtNdl0ZJPbUkbfMGgRuI0f03f9GGfQFbVlIbK1h3E4/838RDbhhRIPjV1v4h2ER2KiG6BcZHDxLlmWgXwAAAABJRU5ErkJggg==";
                }
                else if (d.Type === 'Regulatory') {
                    icon = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAB4AAAAdCAMAAAH9fRyoAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAC9UExURQAAAFlZWVlZWVhYWFtbW1lZWVhYWFlZWVhYWFhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFpaWlhYWFhYWFpaWlhYWFlZWVlZWVlZWVdXV1hYWFlZWVdXV1hYWFhYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVhYWFlZWVpaWldXV1hYWFlZWVhYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVlZWVZWVsacr8UAAAA/dFJOUwCLeJMc5++AmzQk96MEPP+rs0S7VEzDy2TTdNtsIIfjfAiP6ygQl5/zMECn+684SLcMUL/HWM9oYNcUcIPfLAaoF3QAAAAJcEhZcwAAFxEAABcRAcom8z8AAAFuSURBVChTrZNbV4JAFIW3V0zFg6akECmZRmoohRei6f//rM4MM8Yy1+ql7wHO3pszA8MMtmDIsiwQqlw1pZQeW49AjIDrNCg8oG0K6dxVcqFypC15Vc0leAiJVpIXUbTXE+zl/YBeTd7tkOeWQ9t93WRZRzOV5vwOGqMP9i7/0c0NZoNeVjH69iuuCgj+zABvRDuEw8zjgGiOjKgnOw0nwNeTM6HUJTYXmvCqnyygX/lVfXLhJyVNR1GPz3qY5kkLDWF03OQ1erenRj9F6Iy90Df63hly6WdK+8Qf1hyI+g3WRGP2mNME4qMo/VrAVKaLQl5C6n2wfb6CuzPxVQICufrRKzgc/9H9T/GIPjstqzOLPG2UYvuhapHa7+2olrfUxi3FacrLlLiiTyMg7M6VWRo8oBWWXWAwUaX2TLw71Pk6nDeWNrDoFyeBY/Eg13wROrXuhsflrR853jpIpLuXv9WwIm/anW28SJ4EBfANCRcrEWepC3sAAAAASUVORK5CYII=";
                }
                else if (d.Type === 'Launch') {
                    icon = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACEAAAAgCAMAAAHcYdFuAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADAUExURQAAAFlZWVlZWVhYWFtbW1lZWVhYWFlZWVhYWFhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFpaWlhYWFhYWFpaWlhYWFhYWFVVVVlZWVlZWVlZWVdXV1hYWFlZWVdXV1hYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVhYWFlZWVpaWldXV1hYWFlZWVhYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVlZWVZWVsg7ePwAAABAdFJOUwCLeJMc5++AmzQk96MEPP+rs0S7VEzDXBjLZNN022wg43wIj+soEJef8zBAp/uvOEi3DFC/x1jPaGDXFHCD3ywFm3z4AAAACXBIWXMAABcRAAAXEQHKJvM/AAACAElEQVQoU2WTi3aiMBCGfy9oLaVhdxUvqHURL6hYpVYXltX3f6v9JwSt9juHycwkzGQmCYQBP3sPjBUoUT9pL/4BqQscRM+pFNjAAohQoR6KI3MDR0Z0ACWhDHW/GJ+1bK0pYnDlO0Qezv06UplJMiiZRNBgAI6jjPnE8YXU3gGWMYjCmal2M9G9pxZOCdrLGOsug8uy1m+Zwfz6y3A18wCfNZX48CoToGrMacIMEnhx0fYV1TbKPUeX27wRbhyMjU7GumtaaPJ8cOYg5XFXu86xik9s0JMOzKJxkLDnatnWC1xxRrIQ2P6lLVWy7YIn+vLgWyqBqFGi3cSGleBiFRmId1EOnqYmJXlJenb4JnVrghN+xAnb9M4DEfp6u8xtUmG4n1P+DOLCLFk92N9w2vs46i+HFX09vjE72FUJTGq3Bt3IVXkkJHy8CEBX5eio7Vbp83JvIQZerDQvNNxf/OTS1hqBnnWq9p7Vk61V12Mehsw0nxRX23EtfZsYRx0LRZMcGq9aWSpv2mym6UfH2ojd1F6sRlkR9kYt7LFMqyZ6sFA8hwd6WTxaW67406hfX1/PqKTFF9HL/BWeT40P2t0vV98wP3++vU78xrlogDfVg6HnjfgSGP2PCPk5MetKsi4c35SAlBvI76omaWhvTEeAtZrGD++g/Nkwn9zPA/8Bbo4sLEo4ZKcAAAAASUVORK5CYII=";
                }
                return icon;
            })
            .attr('x', '60')
            .attr('y', (d, i) => { if (i % 2 === 0) { return -50; } else { return -35; } })
            .attr('fill', '#000000').attr('transform', 'translate(0,20)');

        this.tooltip
            .style("opacity", 0).attr("class", "tooltip").style("position", "absolute")
            .style("background-color", "white").style("border", "solid").style("border-width", "2px")
            .style("border-radius", "5px").style("padding", "5px");

        let self = this;

        gbox.on('mouseover', (d) => {
            self.tooltip.style("opacity", 1);
        })
            .on('mousemove', (d: TimelineData) => {
                var html = '';
                var dateHtml = '<div>' + (d.Date.getMonth() + 1) + '/' + (d.Date.getDate()) + '/' + (d.Date.getFullYear()) + '</div>';
                if (d.Type === "Regulatory") {
                    html = dateHtml + '<div>' + sanitizeHtml(d.Company) + '</div>' + '<div>' + sanitizeHtml(d.Description) + '</div>';
                }
                else if (d.Type === "Commercial") {
                    html = dateHtml + '<div>' + sanitizeHtml(d.Company) + '</div>' + '<div>' + sanitizeHtml(d.Description) + '</div>';
                } else if (d.Type === "Clinical Trials") {
                    html = dateHtml + '<div>' + sanitizeHtml(d.Company) + '</div>' + '<div>' + + sanitizeHtml(d.Description) + '</div>';
                } else if (d.Type === "Launch") {
                    html = dateHtml + '<div>' + sanitizeHtml(d.Company) + '</div>' + '<div>' + + sanitizeHtml(d.Description) + '</div>';
                }
                self.tooltip.html(html).style("left", (d3.event.pageX + 20) + "px").style("top", (d3.event.pageY) + "px");
            })
            .on('mouseleave', (d) => { self.tooltip.style("opacity", 0); });

        gbox.on("mouseenter", function () { d3.select(this).raise(); });

        this.renderLegend(gWidth, gHeight);
    }

    private renderLegend(gWidth, gHeight) {
        var gLegend = this.svg.append('g')
            .attr('transform', 'translate(' + ((gWidth / 2) - 200) + ',' + (gHeight + 45) + ')')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', gWidth)
            .attr('height', 50);

        var legendClinical = gLegend.append('g')
            .attr('transform', 'translate(50,0)');

        var legendRegulatory = gLegend.append('g')
            .attr('transform', 'translate(200,0)');

        var legendCommercial = gLegend.append('g')
            .attr('transform', 'translate(350,0)');

        var legendLaunch = gLegend.append('g')
            .attr('transform', 'translate(500,0)');

        legendClinical.append("svg:image")
            .attr("xlink:href", "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABoAAAAiCAMAAAEBugmqAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAACoUExURQAAAFlZWVlZWVhYWFlZWVhYWFlZWVhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFhYWFhYWFpaWlVVVVlZWVlZWVlZWVhYWFlZWVdXV1hYWFhYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVhYWFlZWVpaWldXV1hYWFlZWVhYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVZWVrcVYpkAAAA4dFJOUwCLeJPn74A0JPejBDz/q7O7VEwYy2TT22wgh+N8CI/rKBCXn/MwQKf7rzhItwxQv8fPYNcUcN8sDjo1cQAAAAlwSFlzAAAXEQAAFxEByibzPwAAAU1JREFUKFOdku1WwjAMhoMwZEJEvhxzMqabQ5EhLTDv/85M0jBg7ofH55yleZe3bU5bEBCgSwEpIIW2tT5A5GpwpI8M1lpY5ZRbj+XQTWHOI0CRUjiSHrH1GuzryMHQTqKYmBan5aEIH0oDSen3QrH1yIRozls4VI0QFie1xHF/s6ak0anUlZBGB5dAGPrlAEw5DQuALB9RQwTm8xV7d9JLwn3pTKFRbUWcFN5cqO02e/o6KToCHx87lROe+bAr5VAnM22sVdSVL8cqtF71bzFL070Xx50gTWcv+vNw/87Lj1USn7LfcAKZTDTRURey8yDge4wzfi7CIVT0ILvykhr5a6kf4JumxFXJmI8NJMuBU1XpO893Bm5xHZU4o1u6KFF/Exo29JAWrsd6G8meDUK9dEFV4rdbQXfyj1krLXl3nPzm1YOi1VHzFct28QOudyFsm5wOugAAAABJRU5ErkJggg==")
            .attr('transform', 'translate(0,10)');

        legendClinical.append('text')
            .text('Clinical Trials')
            .attr('transform', 'translate(35,35)');

        legendRegulatory.append("svg:image")
            .attr("xlink:href", "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAB4AAAAdCAMAAAH9fRyoAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAC9UExURQAAAFlZWVlZWVhYWFtbW1lZWVhYWFlZWVhYWFhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFpaWlhYWFhYWFpaWlhYWFlZWVlZWVlZWVdXV1hYWFlZWVdXV1hYWFhYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVhYWFlZWVpaWldXV1hYWFlZWVhYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVlZWVZWVsacr8UAAAA/dFJOUwCLeJMc5++AmzQk96MEPP+rs0S7VEzDy2TTdNtsIIfjfAiP6ygQl5/zMECn+684SLcMUL/HWM9oYNcUcIPfLAaoF3QAAAAJcEhZcwAAFxEAABcRAcom8z8AAAFuSURBVChTrZNbV4JAFIW3V0zFg6akECmZRmoohRei6f//rM4MM8Yy1+ql7wHO3pszA8MMtmDIsiwQqlw1pZQeW49AjIDrNCg8oG0K6dxVcqFypC15Vc0leAiJVpIXUbTXE+zl/YBeTd7tkOeWQ9t93WRZRzOV5vwOGqMP9i7/0c0NZoNeVjH69iuuCgj+zABvRDuEw8zjgGiOjKgnOw0nwNeTM6HUJTYXmvCqnyygX/lVfXLhJyVNR1GPz3qY5kkLDWF03OQ1erenRj9F6Iy90Df63hly6WdK+8Qf1hyI+g3WRGP2mNME4qMo/VrAVKaLQl5C6n2wfb6CuzPxVQICufrRKzgc/9H9T/GIPjstqzOLPG2UYvuhapHa7+2olrfUxi3FacrLlLiiTyMg7M6VWRo8oBWWXWAwUaX2TLw71Pk6nDeWNrDoFyeBY/Eg13wROrXuhsflrR853jpIpLuXv9WwIm/anW28SJ4EBfANCRcrEWepC3sAAAAASUVORK5CYII=")
            .attr('transform', 'translate(0, 12)');

        legendRegulatory.append('text')
            .text('Regulatory')
            .attr('transform', 'translate(35,35)');

        legendCommercial.append("svg:image")
            .attr("xlink:href", "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACIAAAAYCAMAAAHff++tAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADYUExURQAAAFlZWVlZWVhYWFtbW1lZWVhYWFlZWVhYWFhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFhYWFhYWFpaWlhYWFhYWFpaWllZWVhYWFhYWFhYWFVVVVlZWVlZWVlZWVdXV1hYWFlZWVdXV1hYWFhYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVlZWVhYWFlZWVpaWldXV1hYWFlZWT8/P1hYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFhYWFlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVlZWVZWVmNebiYAAABIdFJOUwCLeJMc5++AmzQk96MEPP+rs/ThRLtUTOnD8VwYy2TTdNtsIIfjfAiP6ygQl8Wf8zBAp/sIrzhItwxQv9rHWM9oYNcUcIPfLHjl7rgAAAAJcEhZcwAAFxEAABcRAcom8z8AAAGySURBVChTbZGLepNAEIUnVZKadjlqDbGa5loaSBUSVkoo5tYS8v5v5MyySj71/4CdOXNghl26JaIZAXLRdZpzKgAn0O4tNEfI+DYO5jIVIyU3hIpXZdwMl2Kpc1BcFSNrT7FiN0rd9eWtMeDPRYATFdVm4fnJkuWR/bargS3fQq38eCLO9/dGoeDkuG1gjWgF/3U951f3dwhB8zbbqSdzGK73n/pBq469evlDytZeMqoSSv2S86ffvWLXLAajXHk87tTkVmHHYebPjFAr6Tt1AWjVKAbOc54UdHLuh7xjWmMe4lQ8ywS4cahE/oLlGgcRqGJrTJvsiI4nMzHtL9DhI6IlXmtBcGinfBvzbKXswP9R4/hsHEZ7i8KGzvug8+aDCc8sseuaA/ybxhL0zdlc5urRdYxiaSzbB6BaUsb/Hmd53cNw1qh3WA3NKNKupQ72wEEbfsQPz51mDD2lKZfvBosJZ2F9on2OiPJbViZKOm0G3RfpZ6qGEPgJfIycMmtNdl0ZJPbUkbfMGgRuI0f03f9GGfQFbVlIbK1h3E4/838RDbhhRIPjV1v4h2ER2KiG6BcZHDxLlmWgXwAAAABJRU5ErkJggg==")
            .attr('transform', 'translate(0,16)');

        legendCommercial.append('text')
            .text('Commercial')
            .attr('transform', 'translate(35,35)');

        legendLaunch.append("svg:image")
            .attr("xlink:href", "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACEAAAAgCAMAAAHcYdFuAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAADAUExURQAAAFlZWVlZWVhYWFtbW1lZWVhYWFlZWVhYWFhYWFxcXFhYWFlZWT8/P1lZWVlZWVlZWVhYWFpaWlhYWFhYWFpaWlhYWFhYWFVVVVlZWVlZWVlZWVdXV1hYWFlZWVdXV1hYWFhYWF9fX1lZWVhYWFlZWV9fX1lZWVhYWFlZWVpaWldXV1hYWFlZWVhYWFtbW1hYWFlZWVVVVVlZWVlZWVhYWFlZWVhYWFhYWFpaWlhYWFlZWVhYWFlZWVlZWVZWVsg7ePwAAABAdFJOUwCLeJMc5++AmzQk96MEPP+rs0S7VEzDXBjLZNN022wg43wIj+soEJef8zBAp/uvOEi3DFC/x1jPaGDXFHCD3ywFm3z4AAAACXBIWXMAABcRAAAXEQHKJvM/AAACAElEQVQoU2WTi3aiMBCGfy9oLaVhdxUvqHURL6hYpVYXltX3f6v9JwSt9juHycwkzGQmCYQBP3sPjBUoUT9pL/4BqQscRM+pFNjAAohQoR6KI3MDR0Z0ACWhDHW/GJ+1bK0pYnDlO0Qezv06UplJMiiZRNBgAI6jjPnE8YXU3gGWMYjCmal2M9G9pxZOCdrLGOsug8uy1m+Zwfz6y3A18wCfNZX48CoToGrMacIMEnhx0fYV1TbKPUeX27wRbhyMjU7GumtaaPJ8cOYg5XFXu86xik9s0JMOzKJxkLDnatnWC1xxRrIQ2P6lLVWy7YIn+vLgWyqBqFGi3cSGleBiFRmId1EOnqYmJXlJenb4JnVrghN+xAnb9M4DEfp6u8xtUmG4n1P+DOLCLFk92N9w2vs46i+HFX09vjE72FUJTGq3Bt3IVXkkJHy8CEBX5eio7Vbp83JvIQZerDQvNNxf/OTS1hqBnnWq9p7Vk61V12Mehsw0nxRX23EtfZsYRx0LRZMcGq9aWSpv2mym6UfH2ojd1F6sRlkR9kYt7LFMqyZ6sFA8hwd6WTxaW67406hfX1/PqKTFF9HL/BWeT40P2t0vV98wP3++vU78xrlogDfVg6HnjfgSGP2PCPk5MetKsi4c35SAlBvI76omaWhvTEeAtZrGD++g/Nkwn9zPA/8Bbo4sLEo4ZKcAAAAASUVORK5CYII=")
            .attr('transform', 'translate(0,10)');

        legendLaunch.append('text')
            .text('Launch')
            .attr('transform', 'translate(35,35)');
    }

    private renderTimeRangeLines(gHeight, timelineData: TimelineData[]) {
        let _self = this;
        this.svg.selectAll(".line")
            .data(timelineData)
            .enter()
            .append("line")
            .attr("x1", (d: any, i) => {
                return isNaN(_self.xScale(d.Date)) ? 0 : (_self.xScale(d.Date) + 20);
            })
            .attr('y1', (d, i) => {
                if (i % 2 === 0) {
                    return _self.yScale(-10);
                } else {
                    return _self.yScale(10);
                }
            })
            .attr("x2", (d, i) => {
                return isNaN(_self.xScale(d.Date)) ? 0 : (_self.xScale(d.Date) + 20);
            })
            .attr("y2", (d, i) => {
                if (i % 2 === 0) {
                    let count = i / 2;
                    if (count % 2 === 0) {
                        return _self.yScale(-100);
                    }
                    else {
                        return _self.yScale(-60);
                    }
                } else {
                    let count = Math.ceil(i / 2);
                    if (count % 2 === 0) {
                        return _self.yScale(90);
                    }
                    else {
                        return _self.yScale(40);
                    }
                }
            })
            .style('stroke-dasharray', '5,5')
            .style('stroke-width', 2)
            .style('stroke', (d, i) => {
                var colorData = _self.colorDataByYear.find(c => c.year === d.Date.getFullYear());
                return colorData.color;
            })
            .style('fill', 'none');
    }

    private renderXAxisCirclesAndQuarters() {
        let _self = this;
        var tickLength = this.svg.selectAll('.x-axis-line .tick').size();
        this.svg.selectAll('.x-axis-line .tick').insert('circle')
            .attr('cx', 0)
            .attr('cy', 0)
            .attr('r', (d, i) => {
                if (i === 0 || i === (tickLength - 1)) {
                    return 27;
                }
                else {
                    return 10;
                }
            })
            .attr('stroke', (d: Date, i) => {
                if (i === 0 || i === (tickLength - 1)) {
                    return '#868686';
                }
                else {
                    var colorData = _self.colorDataByYear.find(c => c.year === d.getFullYear());
                    return colorData.color;
                }
            })
            .attr('stroke-width', 3)
            .attr('fill', (d: Date, i) => {
                if (i === 0 || i === (tickLength - 1)) {
                    return '#bfbfbf';
                }
                else {
                    var colorData = _self.colorDataByYear.find(c => c.year === d.getFullYear());
                    return colorData.color;
                }
            });
        this.svg.selectAll('.x-axis-line .tick text')
            .attr('y', (d, i) => {
                if (i === 0 || i === (tickLength - 1)) {
                    return 0;
                }
                else {
                    return -30;
                }
            })
            .attr('fill', (d: Date, i) => {
                if (i === 0 || i === (tickLength - 1)) {
                    return '#bfbfbf';
                }
                else {
                    var colorData = _self.colorDataByYear.find(c => c.year === d.getFullYear());
                    return colorData.color;
                }
            }).raise();
    }

    private renderTitle(vpWidth, gWidth) {
        var gTitle = this.svg.append('g')
            .attr('x', 0)
            .attr('y', 0)
            .attr('width', vpWidth)
            .attr('height', 35);

        gTitle.append('rect')
            .attr('class', 'chart-header')
            .attr('width', vpWidth)
            .attr('height', 35);

        gTitle.append('text')
            .text('Key Events Timeline')
            .attr('fill', '#ffffff')
            .attr('font-size', 24)
            .attr('transform', 'translate(' + ((gWidth + 70) / 2 - 104) + ',25)');
    }

    private renderXandYAxis(minDate, maxDate, gWidth, gHeight) {
        var xAxis;
        if (this.diff_years(minDate, maxDate) <= 1) {
            var minDateMonthUpdated = new Date(minDate.getFullYear(), minDate.getMonth() - 1, 1);
            var maxDateMonthUpdated = new Date(maxDate.getFullYear(), maxDate.getMonth() + 1, 1);
            this.xScale = d3.scaleTime()
                .domain([minDateMonthUpdated, maxDateMonthUpdated])//.nice()
                .range([this.margin.left, gWidth]);

            xAxis = d3.axisBottom(this.xScale)
                .ticks(d3.timeMonth, 1)
                .tickPadding(20)
                .tickFormat(d3.timeFormat("%b '%y"))
                .tickSize(-10);
        }
        else {

            this.xScale = d3.scaleTime()
                .domain([minDate, maxDate])
                .range([this.margin.left, gWidth]);

            xAxis = d3.axisBottom(this.xScale)
                .ticks(d3.timeYear, 1)
                .tickPadding(20)
                .tickFormat(d3.timeFormat('%Y'))
                .tickSize(-10);
        }

        var xAxisAllTicks = d3.axisBottom(this.xScale)
            .ticks(d3.timeMonth, 3)
            .tickPadding(20)
            .tickFormat(d3.timeFormat(""))
            .tickSize(10);

        this.yScale = d3.scaleLinear()
            .domain([-105, 105])
            .range([gHeight, this.margin.top]);

        var yAxis = d3.axisLeft(this.yScale);

        var xAxisLine = this.svg.append("g")
            .attr("class", "x-axis-line")
            .attr("transform", "translate(" + (20) + "," + ((gHeight / 2) + 25) + ")")
            .call(xAxis);

        this.svg.append("g")
            .attr("class", "y-axis")
            .call(yAxis).attr('display', 'none');
    }

    private getColorDataByYear(minDate, maxDate) {
        var colors = ['#242B47', '#D5792E', '#8EB40E', '#597DAB', '#5AC1C4', '#595959', '#154360', '#0B5345', '#784212', '#424949',
            '#17202A', '#E74C3C', '#00ff00', '#0000ff', '#252D48'];
        this.colorDataByYear = [];
        for (var year = minDate.getFullYear(), i = 0; year <= maxDate.getFullYear() + 1; year++) {
            this.colorDataByYear.push({
                year: year,
                color: colors[i++]
            });
        }
    }

    private renderHeaderAndFooter(timelineData: TimelineData[]) {
        let [timeline] = timelineData;
        if (this.settings.dataPoint.layout.toLowerCase() === "header") {
            this.header
                .attr("class", "header")
                .append("img")
                .attr(
                    "src",
                    validDataUrl(timeline.HeaderImage) ? timeline.HeaderImage : ""
                ).exit().remove();
        } else if (this.settings.dataPoint.layout.toLowerCase() === "footer") {
            this.footer
                .attr("class", "footer")
                .append("img")
                .attr(
                    "src",
                    validDataUrl(timeline.FooterImage) ? timeline.FooterImage : ""
                );
        }
    }

    public static CONVERTER(dataView: DataView, host: IVisualHost): TimelineData[] {
        var resultData: TimelineData[] = [];
        var tableView = dataView.table;
        var _rows = tableView.rows;
        var _columns = tableView.columns;
        var _companyIndex = -1, _companylinkIndex = -1, _typeIndex = -1, _descIndex = -1, _dateIndex = -1, _headerImageIndex = -1, _footerImageIndex = -1;
        for (var ti = 0; ti < _columns.length; ti++) {
            if (_columns[ti].roles.hasOwnProperty("Company")) {
                _companyIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Type")) {
                _typeIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Description")) {
                _descIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("CompanyLink")) {
                _companylinkIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("Date")) {
                _dateIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("HeaderImage")) {
                _headerImageIndex = ti;
            } else if (_columns[ti].roles.hasOwnProperty("FooterImage")) {
                _footerImageIndex = ti;
            }
        }
        for (var i = 0; i < _rows.length; i++) {
            var row = _rows[i];
            var dp = {
                Company: row[_companyIndex].toString(),
                Type: row[_typeIndex] ? row[_typeIndex].toString() : '',
                Description: row[_descIndex] ? row[_descIndex].toString() : null,
                CompanyLink: row[_companylinkIndex] ? row[_companylinkIndex].toString() : null,
                Date: row[_dateIndex] ? new Date(Date.parse(row[_dateIndex].toString())) : null,
                HeaderImage: row[_headerImageIndex] ? row[_headerImageIndex].toString() : null,
                FooterImage: row[_footerImageIndex] ? row[_footerImageIndex].toString() : null,
                selectionId: host.createSelectionIdBuilder().withTable(tableView, i).createSelectionId(),
            };
            resultData.push(dp);
        }
        return resultData;
    }

    private static parseSettings(dataView: DataView): VisualSettings {
        return VisualSettings.parse(dataView);
    }

    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstance[] | VisualObjectInstanceEnumerationObject {
        return VisualSettings.enumerateObjectInstances(this.settings || VisualSettings.getDefault(), options);
    }

    private wrap(text, width) {
        text.each(function () {
            var text = d3.select(this),
                words = text.text().split(/\s+/).reverse(),
                word,
                line = [],
                lineNumber = 0,
                lineHeight = 1.1,
                x = text.attr("x"),
                y = text.attr("y"),
                dy = 0,
                tspan = text.text(null)
                    .append("tspan")
                    .attr("x", x)
                    .attr("y", y)
                    .attr("dy", dy + "em");
            while (word = words.pop()) {
                line.push(word);
                tspan.text(line.join(" "));
                if (tspan.node().getComputedTextLength() > width) {
                    line.pop();
                    tspan.text(line.join(" "));
                    line = [word];
                    tspan = text.append("tspan")
                        .attr("x", x)
                        .attr("y", y)
                        .attr("dy", ++lineNumber * lineHeight + dy + "em")
                        .text(word);
                }
            }
        });
    }

    private diff_years(dt2, dt1) {
        var diff = (dt2.getTime() - dt1.getTime()) / 1000;
        diff /= (60 * 60 * 24);
        return Math.abs(Math.round(diff / 365.25));
    }
}