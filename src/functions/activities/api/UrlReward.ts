import { randomUUID } from 'node:crypto'
import { URLs } from '../../../constants/urls'
import { BING_APP_USER_AGENT } from '../../../constants/userAgents'
import type { BasePromotion } from '../../../interface/DashboardData'
import { BaseActivity } from '../BaseActivity'

export class UrlReward extends BaseActivity {
    public async doUrlReward(promotion: BasePromotion) {
        await this.runUrlReward(promotion, true)
    }

    private async runUrlReward(promotion: BasePromotion, allowSessionRepair: boolean) {
        const offerId = promotion.offerId

        const actionId = this.bot.nextActions.reportActivity
        if (!actionId) {
            this.bot.logger.warn(
                this.bot.isMobile,
                'URL-REWARD',
                `跳过 ${offerId}：bundle 中未发现 "reportActivity"`
            )
            return
        }

        const live = await this.bot.browser.func.ensureOffer(offerId)
        if (!live) {
            // 页面快照缺失：回退到 App dapi 官方通道上报（V4-china / geo=CN 下每日任务在页面快照中无 offerId）
            if (await this.dapiReportActivity(promotion)) return
            this.bot.logger.warn(
                this.bot.isMobile,
                'URL-REWARD',
                `跳过 ${offerId}：页面快照中不存在（dapi 上报亦未成功），即使重新获取 /earn 和 /dashboard 后也没有`
            )
            return
        }
        if (!live.reportable) {
            this.bot.logger.warn(
                this.bot.isMobile,
                'URL-REWARD',
                `跳过 ${offerId}：不可上报（已完成/已锁定/无 hash/未到期）`
            )
            return
        }

        if (this.bot.config.skipNonPointTasks && live.points === 0) {
            this.bot.logger.info(
                this.bot.isMobile,
                'URL-REWARD',
                `跳过 ${offerId}：不产生积分（points=${live.points}${live.promotionSubtype ? ` subtype=${live.promotionSubtype}` : ''}）- 可能是免费试用/不计分优惠。设置 skipNonPointTasks=false 可强制尝试。`
            )
            return
        }

        const oldBalance = this.bot.userData.currentPoints
        const expectedPoints = live.points

        const dashboardActivityType = Number(promotion.activityType)
        const activityType =
            live.activityType ??
            (Number.isInteger(dashboardActivityType) && dashboardActivityType > 0 ? dashboardActivityType : 11)

        this.bot.logger.info(
            this.bot.isMobile,
            'URL-REWARD',
            `开始处理 UrlReward | offerId=${offerId} | geo=${this.bot.userData.geoLocale} | currentBalance=${oldBalance}`
        )

        try {
            const { status, acknowledged, availablePoints } = await this.bot.browser.func.reportServerAction(
                actionId,
                [
                    live.hash,
                    activityType,
                    {
                        offerid: offerId,
                        isPromotional: live.isPromotional ? true : '$undefined',
                        timezoneOffset: this.bot.userData.timezoneOffset
                    }
                ],
                {
                    url: URLs.rewards.dashboard,
                    referer: URLs.rewards.dashboard,
                    routerStateTree: this.bot.browser.react.routerStateTree('dashboard')
                }
            )

            if (!acknowledged) {
                this.bot.logger.warn(
                    this.bot.isMobile,
                    'URL-REWARD',
                    `UrlReward 请求未被确认 | offerId=${offerId} | status=${status}`
                )
                if (await this.retryAfterRequestFailure(promotion, allowSessionRepair)) return
            }

            const newBalance = availablePoints ?? (await this.bot.browser.func.getCurrentPoints())
            const gainedPoints = newBalance - oldBalance

            this.bot.logger.debug(
                this.bot.isMobile,
                'URL-REWARD',
                `响应 | offerId=${offerId} | status=${status} | acknowledged=${acknowledged} | pointsGained=${gainedPoints} | currentBalance=${newBalance}`
            )

            if (gainedPoints > 0) {
                this.bot.userData.currentPoints = newBalance
                this.bot.userData.gainedPoints = (this.bot.userData.gainedPoints ?? 0) + gainedPoints

                const shortfall = expectedPoints > 0 && gainedPoints < expectedPoints
                this.bot.logger.info(
                    this.bot.isMobile,
                    'URL-REWARD',
                    `UrlReward 完成 | offerId=${offerId} | pointsGained=${gainedPoints} | currentBalance=${newBalance}${shortfall ? ' | WARNING: credited less than advertised' : ''}`,
                    'green'
                )
            } else if (acknowledged && expectedPoints === 0) {
                this.bot.logger.info(
                    this.bot.isMobile,
                    'URL-REWARD',
                    `UrlReward 完成（设计上无积分） | offerId=${offerId} | acknowledged=true | pointsGained=0 | currentBalance=${newBalance}`,
                    'green'
                )
            } else {
                this.bot.logger.warn(
                    this.bot.isMobile,
                    'URL-REWARD',
                    `UrlReward 未获得积分 | offerId=${offerId} | acknowledged=${acknowledged} | expected=${expectedPoints} | pointsGained=0 | currentBalance=${newBalance}`
                )
            }
        } catch (error) {
            this.bot.logger.error(
                this.bot.isMobile,
                'URL-REWARD',
                `doUrlReward 出错 | offerId=${offerId} | message=${error instanceof Error ? error.message : String(error)}`
            )
            await this.retryAfterRequestFailure(promotion, allowSessionRepair)
        }
    }

    private async retryAfterRequestFailure(promotion: BasePromotion, allowSessionRepair: boolean): Promise<boolean> {
        if (!allowSessionRepair) return false

        const refreshed = await this.bot.refreshCurrentRewardsContext(`URL-REWARD:${promotion.offerId}`)
        if (!refreshed) return false

        this.bot.logger.info(
            this.bot.isMobile,
            'URL-REWARD',
            `使用刷新后的 cookie 和引导数据重试一次 UrlReward | offerId=${promotion.offerId}`
        )
        await this.runUrlReward(promotion, false)
        return true
    }
    /**
     * App dapi 通道上报（Microsoft Rewards 每日任务官方上报接口）：
     * POST prod.rewardsplatform.microsoft.com/dapi/me/activities
     * 使用登录时获取的 MBI_SSL accessToken（this.bot.accessToken）。
     * 请求格式对照 BingSapphire App：attributes.offerid + amount + type:101 + channel + id + meta + country。
     * meta 为 App 端签名，实测服务端不校验（空串可接受）。
     */
    private async dapiReportActivity(promotion: BasePromotion): Promise<boolean> {
        const offerId = promotion.offerId
        const token = this.bot.accessToken
        if (!token) {
            this.bot.logger.warn(this.bot.isMobile, 'URL-REWARD', `dapi 上报跳过 ${offerId}：无 accessToken（未走移动端登录）`)
            return false
        }

        this.bot.logger.info(
            this.bot.isMobile,
            'URL-REWARD',
            `dapi 通道上报 | offerId=${offerId} | type=101 | currentBalance=${this.bot.userData.currentPoints}`
        )

        try {
            const oldBalance = this.bot.userData.currentPoints
            const request = {
                url: URLs.platform.activities,
                method: 'POST' as const,
                headers: {
                    'Content-Type': 'application/json',
                    Authorization: `Bearer ${token}`,
                    'User-Agent': BING_APP_USER_AGENT,
                    'X-Rewards-Country': this.bot.userData.geoLocale,
                    'X-Rewards-Language': this.bot.userData.langCode,
                    'X-Rewards-IsMobile': 'true',
                    Accept: '*/*'
                },
                data: {
                    attributes: { offerid: offerId },
                    amount: 1,
                    type: 101,
                    risk_context: {},
                    channel: 'SAAndroid',
                    id: randomUUID(),
                    meta: '',
                    country: 'cn'
                },
                retries: 1
            }

            const response = await this.bot.http.request<{ code?: number; response?: { balance?: number } }>(
                request,
                false
            )
            const code = response.data?.code
            const newBalance = response.data?.response?.balance ?? (await this.bot.browser.func.getCurrentPoints())
            const gainedPoints = newBalance - oldBalance

            this.bot.logger.info(
                this.bot.isMobile,
                'URL-REWARD',
                `dapi 上报响应 | offerId=${offerId} | http=${response.status} | code=${code} | pointsGained=${gainedPoints} | balance=${newBalance} | raw=${JSON.stringify(response.data ?? '').slice(0, 600)}`
            )

            if (code === 0) {
                if (gainedPoints > 0) {
                    this.bot.userData.currentPoints = newBalance
                    this.bot.userData.gainedPoints = (this.bot.userData.gainedPoints ?? 0) + gainedPoints
                }
                this.bot.logger.info(
                    this.bot.isMobile,
                    'URL-REWARD',
                    `dapi 上报完成 | offerId=${offerId} | pointsGained=${gainedPoints}`,
                    'green'
                )
                return true
            }

            this.bot.logger.warn(
                this.bot.isMobile,
                'URL-REWARD',
                `dapi 上报未成功 | offerId=${offerId} | http=${response.status} | code=${code} | raw=${JSON.stringify(response.data ?? '').slice(0, 600)}`
            )
            return false
        } catch (error) {
            this.bot.logger.error(
                this.bot.isMobile,
                'URL-REWARD',
                `dapi 上报出错 | offerId=${offerId} | message=${error instanceof Error ? error.message : String(error)}`
            )
            return false
        }
    }


}
