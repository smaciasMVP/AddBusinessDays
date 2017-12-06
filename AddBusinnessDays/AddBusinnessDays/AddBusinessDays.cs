// <copyright file="AddBusinessDays.cs" company="">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author></author>
// <date>4/21/2017 3:24:54 PM</date>
// <summary>Implements the AddBusinessDays Workflow Activity.</summary>
namespace AddBusinnessDays
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;


    public sealed class AddBusinessDays : CodeActivity
    {
        [RequiredArgument]
        [Input("Original Date")]
        public InArgument<DateTime> OriginalDate { get; set; }

        [RequiredArgument]
        [Input("Business Days To Add")]
        public InArgument<int> BusinessDaysToAdd { get; set; }

        [Input("Holiday/Closure Calendar")]
        [ReferenceTarget("calendar")]
        public InArgument<EntityReference> HolidayClosureCalendar { get; set; }

        [OutputAttribute("Updated Date")]
        public OutArgument<DateTime> UpdatedDate { get; set; }

        [RequiredArgument]
        [Input("IsHighLevel")]
        public InArgument<Boolean> IsHighLevel { get; set; }

        [RequiredArgument]
        [Input("DecreaseHours")]
        [Default("0")]
        public InArgument<Int32> DecreaseHours { get; set; }




        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        /// 
        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                DateTime originalDate = OriginalDate.Get(executionContext);
                int businessDaysToAdd = BusinessDaysToAdd.Get(executionContext);
                EntityReference holidaySchedule = HolidayClosureCalendar.Get(executionContext);

                Boolean ishighLevel = IsHighLevel.Get(executionContext);
                Int32 decreaseHours = DecreaseHours.Get(executionContext);

                Entity calendar = null;
                EntityCollection calendarRules = null;
                if (holidaySchedule != null)
                {
                    calendar = service.Retrieve("calendar", holidaySchedule.Id, new ColumnSet(true));
                    if (calendar != null)
                        calendarRules = calendar.GetAttributeValue<EntityCollection>("calendarrules");
                }

                DateTime tempDate = originalDate;
                DateTime tempDateHours = originalDate;

                /*
                 
                 Test segment
                 */
                if (ishighLevel)
                {
                    Int32 DecreaseHoursNegative = decreaseHours * -1;
                    tempDateHours = tempDateHours.AddHours(DecreaseHoursNegative);
                    if (tempDateHours.Day != tempDate.Day)
                    {
                        businessDaysToAdd = -1;
                        tempDate = tempDate.AddHours(DecreaseHoursNegative);
                    }

                }

                /****************/

                if (businessDaysToAdd > 0)
                {
                    while (businessDaysToAdd > 0)
                    {
                        tempDate = tempDate.AddDays(1);
                        if (tempDate.DayOfWeek == DayOfWeek.Sunday || tempDate.DayOfWeek == DayOfWeek.Saturday)
                            continue;

                        if (calendar == null)
                        {
                            businessDaysToAdd--;
                            continue;
                        }

                        bool isHoliday = false;
                        foreach (Entity calendarRule in calendarRules.Entities)
                        {
                            DateTime startTime = calendarRule.GetAttributeValue<DateTime>("starttime");

                            //Not same date
                            if (!startTime.Date.Equals(tempDate.Date))
                                continue;

                            //Not full day event
                            if (startTime.Subtract(startTime.TimeOfDay) != startTime || calendarRule.GetAttributeValue<int>("duration") != 1440)
                                continue;

                            isHoliday = true;
                            break;
                        }
                        if (!isHoliday)
                            businessDaysToAdd--;
                    }
                }
                else if (businessDaysToAdd < 0)
                {
                    while (businessDaysToAdd < 0)
                    {
                        tempDate = tempDate.AddDays(-1);
                        if (tempDate.DayOfWeek == DayOfWeek.Sunday || tempDate.DayOfWeek == DayOfWeek.Saturday)
                            continue;

                        if (calendar == null)
                        {
                            businessDaysToAdd++;
                            continue;
                        }

                        bool isHoliday = false;
                        foreach (Entity calendarRule in calendarRules.Entities)
                        {
                            DateTime startTime = calendarRule.GetAttributeValue<DateTime>("starttime");

                            //Not same date
                            if (!startTime.Date.Equals(tempDate.Date))
                                continue;

                            //Not full day event
                            if (startTime.Subtract(startTime.TimeOfDay) != startTime || calendarRule.GetAttributeValue<int>("duration") != 1440)
                                continue;

                            isHoliday = true;
                            break;
                        }
                        if (!isHoliday)
                            businessDaysToAdd++;
                    }
                }
                

                DateTime updatedDate = tempDate;

                UpdatedDate.Set(executionContext, updatedDate);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        
        }
    }
}